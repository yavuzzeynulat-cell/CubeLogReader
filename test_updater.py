"""Unit tests for updater.py - uses unittest + tempfile, no network."""
import hashlib
import os
import shutil
import sys
import tempfile
import unittest
import zipfile
from unittest.mock import patch

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import updater


def _mock_release(tag="v1.0.1", body="Changes:\n- Bug fix\n\nSHA256: abc123",
                  asset_name="src.zip", asset_url="https://example.com/src.zip"):
    return {
        "tag_name": tag,
        "body": body,
        "assets": [{"name": asset_name, "browser_download_url": asset_url}],
    }


def _make_test_zip(path, files):
    with zipfile.ZipFile(path, "w") as zf:
        for name, content in files.items():
            zf.writestr(name, content)


def _sha256_of(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


class CheckForUpdateTests(unittest.TestCase):

    def setUp(self):
        self._patcher = patch.object(updater, "_read_local_version", return_value="1.0.0")
        self._patcher.start()

    def tearDown(self):
        self._patcher.stop()

    @patch("updater._fetch_latest_release_json")
    def test_returns_update_info_when_remote_newer(self, mock_fetch):
        mock_fetch.return_value = _mock_release(tag="v1.0.1")
        info = updater.check_for_update()
        self.assertIsNotNone(info)
        self.assertEqual(info.version, "1.0.1")
        self.assertEqual(info.asset_url, "https://example.com/src.zip")
        self.assertEqual(info.sha256, "abc123")
        self.assertIn("Bug fix", info.notes)

    @patch("updater._fetch_latest_release_json")
    def test_returns_none_when_same_version(self, mock_fetch):
        mock_fetch.return_value = _mock_release(tag="v1.0.0")
        self.assertIsNone(updater.check_for_update())

    @patch("updater._fetch_latest_release_json")
    def test_returns_none_when_older_remote(self, mock_fetch):
        mock_fetch.return_value = _mock_release(tag="v0.9.0")
        self.assertIsNone(updater.check_for_update())

    @patch("updater._fetch_latest_release_json")
    def test_returns_none_on_network_error(self, mock_fetch):
        mock_fetch.side_effect = OSError("network down")
        self.assertIsNone(updater.check_for_update())

    @patch("updater._fetch_latest_release_json")
    def test_returns_none_on_bad_json(self, mock_fetch):
        mock_fetch.return_value = {"unexpected": "shape"}
        self.assertIsNone(updater.check_for_update())

    @patch("updater._fetch_latest_release_json")
    def test_handles_missing_sha256_line(self, mock_fetch):
        mock_fetch.return_value = _mock_release(body="Just a fix, no checksum")
        info = updater.check_for_update()
        self.assertIsNotNone(info)
        self.assertIsNone(info.sha256)


class DownloadUpdateTests(unittest.TestCase):

    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()
        self.dest = os.path.join(self.tmpdir.name, "src.zip")
        self.remote = os.path.join(self.tmpdir.name, "remote_src.zip")
        _make_test_zip(self.remote, {"main.py": "print('hi')", "version.txt": "1.0.1"})
        self.remote_sha = _sha256_of(self.remote)

    def tearDown(self):
        self.tmpdir.cleanup()

    def _info(self, sha=None):
        return updater.UpdateInfo(
            version="1.0.1", notes="", asset_url="https://example.com/src.zip",
            sha256=sha,
        )

    @patch("updater._http_download")
    def test_download_success_with_correct_sha(self, mock_dl):
        mock_dl.side_effect = lambda url, dest, timeout=30: shutil.copy(self.remote, dest)
        ok = updater.download_update(self._info(sha=self.remote_sha), self.dest)
        self.assertTrue(ok)
        self.assertTrue(os.path.isfile(self.dest))

    @patch("updater._http_download")
    def test_download_succeeds_when_sha_absent(self, mock_dl):
        mock_dl.side_effect = lambda url, dest, timeout=30: shutil.copy(self.remote, dest)
        ok = updater.download_update(self._info(sha=None), self.dest)
        self.assertTrue(ok)

    @patch("updater._http_download")
    def test_download_fails_on_sha_mismatch(self, mock_dl):
        mock_dl.side_effect = lambda url, dest, timeout=30: shutil.copy(self.remote, dest)
        ok = updater.download_update(self._info(sha="0" * 64), self.dest)
        self.assertFalse(ok)
        self.assertFalse(os.path.isfile(self.dest))

    @patch("updater._http_download")
    def test_download_fails_on_corrupt_zip(self, mock_dl):
        def write_garbage(url, dest, timeout=30):
            with open(dest, "wb") as f:
                f.write(b"not a zip")
        mock_dl.side_effect = write_garbage
        ok = updater.download_update(self._info(sha=None), self.dest)
        self.assertFalse(ok)

    @patch("updater._http_download")
    def test_download_fails_on_network_error(self, mock_dl):
        mock_dl.side_effect = OSError("connection refused")
        ok = updater.download_update(self._info(), self.dest)
        self.assertFalse(ok)


class ApplyUpdateTests(unittest.TestCase):

    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()
        self.app_dir = self.tmpdir.name
        self.src = os.path.join(self.app_dir, "src")
        self.backup = os.path.join(self.app_dir, "src_backup")
        os.makedirs(self.src)
        with open(os.path.join(self.src, "main.py"), "w") as f:
            f.write("# old main")
        with open(os.path.join(self.src, "version.txt"), "w") as f:
            f.write("1.0.0")
        self.zip_path = os.path.join(self.app_dir, "update.zip")
        _make_test_zip(self.zip_path, {
            "main.py": "# new main",
            "reader.py": "# new reader",
            "version.txt": "1.0.1",
        })
        self._patcher = patch.object(updater, "_app_dir", return_value=self.app_dir)
        self._patcher.start()

    def tearDown(self):
        self._patcher.stop()
        self.tmpdir.cleanup()

    def test_apply_moves_old_src_to_backup(self):
        updater.apply_update(self.zip_path)
        self.assertTrue(os.path.isdir(self.backup))
        with open(os.path.join(self.backup, "main.py")) as f:
            self.assertEqual(f.read(), "# old main")

    def test_apply_extracts_new_files_into_src(self):
        updater.apply_update(self.zip_path)
        with open(os.path.join(self.src, "main.py")) as f:
            self.assertEqual(f.read(), "# new main")
        with open(os.path.join(self.src, "reader.py")) as f:
            self.assertEqual(f.read(), "# new reader")
        with open(os.path.join(self.src, "version.txt")) as f:
            self.assertEqual(f.read(), "1.0.1")

    def test_apply_removes_pre_existing_backup(self):
        os.makedirs(self.backup)
        with open(os.path.join(self.backup, "stale.txt"), "w") as f:
            f.write("stale")
        updater.apply_update(self.zip_path)
        self.assertFalse(os.path.isfile(os.path.join(self.backup, "stale.txt")))
        self.assertTrue(os.path.isfile(os.path.join(self.backup, "main.py")))

    def test_apply_no_existing_src(self):
        shutil.rmtree(self.src)
        updater.apply_update(self.zip_path)
        self.assertTrue(os.path.isfile(os.path.join(self.src, "main.py")))
        self.assertFalse(os.path.isdir(self.backup))


if __name__ == "__main__":
    unittest.main()
