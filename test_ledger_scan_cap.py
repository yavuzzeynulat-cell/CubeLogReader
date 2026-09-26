"""
test_ledger_scan_cap.py — the ledger scan must reach samples past row 30000.

The 2026 Concrete ledger passed 30000 rows (G26-CON-1236 sits at 29236-29241),
and every sample below the old LEDGER_MAX_SCAN_ROWS=30000 cap came back
"not found" even though its block was there.
"""
import writer


class _Used:
    def __init__(self, last_row):
        self.Row = 1
        self.Rows = type("R", (), {"Count": last_row})()


class _Range:
    def __init__(self, value):
        self.Value = value


class _FakeSheet:
    """Rows 8..last_row of A:L; one sample block at `block_row`."""

    def __init__(self, last_row, block_row, mark, cube_no):
        self.UsedRange = _Used(last_row)
        self._last_row = last_row
        self._block_row = block_row
        self._mark = mark
        self._cube_no = cube_no

    def Range(self, addr):
        end = int(addr.split(":")[1][1:])
        rows = []
        for r in range(8, end + 1):
            row = [None] * 12
            if r == self._block_row:
                row[0] = self._cube_no
                row[1] = self._mark
            if self._block_row <= r < self._block_row + 3:
                row[11] = 7
            rows.append(tuple(row))
        return _Range(tuple(rows))


def test_block_past_row_30000_is_read():
    ws = _FakeSheet(last_row=30800, block_row=30500,
                    mark="G26-CON-1352", cube_no=1265)
    blocks = writer.read_ledger_blocks(ws)
    keys = [b["sample_key"] for b in blocks]
    assert "G26-CON-1352" in keys
    b = blocks[keys.index("G26-CON-1352")]
    assert b["rows_7d"] == [30500, 30501, 30502]


def test_shotcrete_block_past_row_30000_is_read():
    ws = _FakeSheet(last_row=30800, block_row=30500,
                    mark="G26-SHC-0100", cube_no=100)
    # Shotcrete reads A:K, age in K (index 10): move the age marks over.
    orig = ws.Range

    def _range_k(addr):
        rows = tuple(r[:10] + (r[11],) for r in orig(addr).Value)
        return _Range(rows)

    ws.Range = _range_k
    blocks = writer.read_shotcrete_ledger_blocks(ws)
    assert "G26-SHC-100" in [b["sample_key"] for b in blocks]
