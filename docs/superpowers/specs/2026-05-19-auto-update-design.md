# CubeLogReader — Auto-Update Tasarımı

**Tarih:** 2026-05-19
**Yazar:** Yavuz Zeynula (with Claude)
**Durum:** Onay bekliyor (kullanıcı review aşaması)

---

## 1. Amaç

CubeLogReader şirket bilgisayarlarında kullanıldığı için, geliştirici makinesinde
yapılan değişikliklerin **otomatik olarak** ve **antivirus (AV) tarafından
engellenmeden** son kullanıcı bilgisayarına ulaşması gerekir. Tek kullanıcı
şu an Yavuz; ileride başka kullanıcılara genişleyebilir.

### Başarı kriterleri

- Geliştirici tarafında: kod değişikliği → tek komutla GitHub release → bitti.
- Kullanıcı tarafında: uygulama açılışında yeni sürümü kendi bulur, izinle
  güncellenir, restart sonrası yeni kodla çalışır.
- AV taraması: exe **hiç değişmediği için** ilk whitelist sonrası AV bir daha
  tarama tetiklemez. `.exe` indirilmez, çalıştırılmaz.
- **Hiçbir update programı bozamaz** — bozuk sürüm açılışta tespit edilir,
  otomatik bir önceki sürüme geri dönülür.

---

## 2. Mimari Karar: Launcher + harici `src/`

PyInstaller `--onedir` build'i korunur. Tek mimari değişiklik: uygulama
kaynak kodu (`main.py`, `reader.py`, `writer.py`) artık **exe içinde gömülü
değildir**; exe'nin yanında `src/` klasöründe ayrı `.py` dosyaları olarak durur.
Exe sadece bir **launcher**'dır; PyInstaller tüm bağımlılıkları (CustomTkinter,
google-generativeai, pywin32, vs.) launcher exe'sine bundle eder, ama 4
uygulama dosyasını dışarıda tutar.

Update mekanizması yalnızca bu `src/` klasörünü değiştirir.

### Dağıtım klasörü yapısı (kurulum sonrası)

```
CubeLogReader/
├── CubeLogReader.exe                <- launcher (artık küçük, ~10-15 MB)
├── _internal/                       <- PyInstaller bağımlılıkları (değişmez)
├── src/                             <- YENİ: uygulama kaynağı (update edilir)
│   ├── main.py
│   ├── reader.py
│   ├── writer.py
│   ├── updater.py                   <- YENİ: update mantığı
│   └── version.txt                  <- "1.0.0"
├── src_backup/                      <- update sırasında otomatik (rollback için)
└── .env                             <- Gemini API key (kullanıcıya ait)
```

---

## 3. Bileşenler

### 3.1 `launcher.py` (PyInstaller bunu build eder)

Görevi:

1. `sys.path`'in başına `src/` ekle.
2. `from main import main; main()` çağır.
3. Hata olursa rollback dene (`src_backup/` → `src/`), bir kez yeniden dene.
4. Yine başarısızsa CTk messagebox ile hata göster.

Sözde kod:

```python
import os, sys, shutil, traceback
APP_DIR = os.path.dirname(sys.executable)
SRC = os.path.join(APP_DIR, "src")
BACKUP = os.path.join(APP_DIR, "src_backup")

def _try_run():
    sys.path.insert(0, SRC)
    from main import main
    main()

def _rollback():
    if os.path.isdir(BACKUP):
        if os.path.isdir(SRC):
            shutil.rmtree(SRC)
        shutil.move(BACKUP, SRC)
        return True
    return False

try:
    _try_run()
except Exception:
    err = traceback.format_exc()
    if _rollback():
        # backup geri yüklendi, yeniden başlat
        os.execv(sys.executable, [sys.executable])
    else:
        _show_error_dialog(err)
```

### 3.2 `src/updater.py` (uygulama kodunda yeni modül)

Public API:

- `check_for_update(timeout=5) -> Optional[UpdateInfo]`
  GitHub releases API'sini sorgular, lokal `version.txt` ile karşılaştırır.
  Yeni sürüm yoksa `None`; varsa `UpdateInfo(version, notes, asset_url, sha256)`.
  İnternet/timeout hatalarında `None` döner (sessiz başarısızlık).

- `download_update(info, dest_path) -> bool`
  Asset URL'inden zip indirir, varsa SHA256 doğrular, `zipfile.testzip()` çalıştırır.

- `apply_update(zip_path) -> None`
  1. `src/` → `src_backup/` rename (eski backup varsa önce silinir).
  2. Zip'i yeni `src/`'ye extract eder.
  3. İşlem atomik değildir ama her adım idempotent — yarıda kalırsa launcher
     bir sonraki açılışta yine src_backup'tan toparlar.

- `restart_app() -> None`
  `subprocess.Popen([sys.executable])` + mevcut process `sys.exit(0)`.

Bağımlılık: sadece `urllib.request`, `zipfile`, `hashlib`, `shutil`, `os`, `sys`,
`subprocess`. Hiçbir 3rd party kütüphane eklemez.

### 3.3 `src/main.py` değişiklikleri

- Açılışta arka planda (thread) `updater.check_for_update()` çağırır.
- Sonuç varsa CTk modal: *"Yeni sürüm 1.0.X mevcut. Şimdi güncellensin mi?"*
  ile *"Sonra"* butonları + release notes preview.
- `SettingsDialog` içine **"Güncellemeleri kontrol et"** butonu eklenir; aynı
  fonksiyonu sync çağırır, sonucu inline gösterir.

### 3.4 Build sistemi değişiklikleri

#### `CubeLogReader.spec` (yeni veya güncel)

- `Analysis(scripts=['launcher.py'], ...)`
- `datas=[]` listesinde **`src/*.py`** ve **`src/version.txt`** **YOK** — bunlar
  exe içine gömülmez, yan yana dağıtılır.
- `launcher.py`'nin başına **`import customtkinter; import google.generativeai;
  import win32com.client; import fitz; import pythoncom`** vb. eklenir ki
  PyInstaller bu ağır bağımlılıkları analiz edip bundle'a alsın. `main.py`,
  `reader.py`, `writer.py`, `updater.py` ise PyInstaller'ın project root'unda
  bulunduğu için tarama sırasında dolaylı olarak görünür; ama bundle dışı
  tutulmaları için `excludes=['main','reader','writer','updater']` eklenir.
  Runtime'da `src/` klasöründen import edilirler.

#### `build_exe.bat`

Build sonunda:

```
mkdir dist\CubeLogReader\src
copy main.py    dist\CubeLogReader\src\
copy reader.py  dist\CubeLogReader\src\
copy writer.py  dist\CubeLogReader\src\
copy updater.py dist\CubeLogReader\src\
copy version.txt dist\CubeLogReader\src\
```

#### `installer.iss` (Inno Setup)

- `[Files]` bölümüne `src\*` eklenir.
- Diğer kurulum davranışı aynı.

---

## 4. Update Akışı (state diyagramı)

```
app start
   |
   v
launcher.try_run()
   |
   +-- success --> running app (background: updater.check)
   |                  |
   |                  +-- new version --> dialog --> [user accepts]
   |                                                      |
   |                                                      v
   |                                       updater.download_update()
   |                                                      |
   |                                                      v
   |                                            updater.apply_update()
   |                                                      |
   |                                                      v
   |                                            "restart now" dialog
   |                                                      |
   |                                                      v
   |                                            updater.restart_app()
   |
   +-- exception --> rollback (src_backup -> src) --> os.execv (retry)
                       |
                       +-- backup yok --> error dialog
```

### Sürüm karşılaştırma

- `version.txt` içeriği: `"1.0.1"` (saf semver, başında `v` yok).
- Karşılaştırma: `tuple(map(int, local.split("."))) < tuple(map(int, remote.split(".")))`.
- Geçersiz formatta lokal version → "0.0.0" varsay → her zaman güncelle.

### Release tag formatı

- GitHub'da tag: `v1.0.1` (başında `v`).
- App `tag_name`'den baştaki `v`'yi soyup karşılaştırır.

---

## 5. AV (antivirus) Stratejisi

| Tedbir | Amaç |
|--------|------|
| Exe **hiç değişmez** | İlk whitelist sonrası AV taraması tetiklenmez |
| `--onedir` bundle (mevcut) | `--onefile`'ın self-extract pattern'i AV'yi tetikler |
| Update sadece `.py` indirir, exe indirmez/çalıştırmaz | "Download+execute binary" AV pattern'i yok |
| HTTPS + GitHub | Bilinen domain, şirket firewall'ları güvenir |
| `urllib` (stdlib), 3rd party HTTP yok | Yeni AV-tetikleyici binary eklenmez |
| Code signing certificate (ileride opsiyonel) | AV neredeyse sıfır şüphe duyar |

---

## 6. Hata yönetimi ve risk azaltma

| Risk | Azaltma |
|------|---------|
| Yeni `src/` bozuk (syntax error, import error) | Launcher `try/except` → otomatik backup rollback |
| İndirme yarım kaldı | SHA256 + `zipfile.testzip()` doğrulama |
| GitHub API rate limit (60/saat unauth) | Son kontrol zamanı kaydedilir, 1 saatten erken tekrar denenmez |
| `src_backup/` da bozuksa | Error dialog + manuel installer link'i |
| Update sırasında Excel açıkken Windows .py'leri lock'lar | `.py` lock olmaz; sadece `.exe`/`.dll` Windows lock'lar. Risk yok. |
| Çevrimdışı / VPN yok | `check_for_update` 5sn timeout → None → app normal açılır |
| Eski sürümden çok yeni sürüme atlanırken format değişiklikleri | Major version bump'larda update notes'a uyarı; ilerideki konu |

---

## 7. Test Stratejisi

### 7.1 Birim testleri (`test_updater.py`)

- `check_for_update`: mock GitHub JSON cevabı → doğru version compare.
- `check_for_update`: timeout / 404 / bozuk JSON → `None` döner, exception fırlatmaz.
- `download_update`: SHA256 mismatch → False döner, dosya silinir.
- `apply_update`: backup dizini önce silinir, src → backup, zip → src.
- Hepsi `tempfile.TemporaryDirectory()` içinde, internet gerekmez.

### 7.2 Manuel end-to-end (ilk release öncesi bir kez)

1. `v1.0.0` build + lokal kur.
2. Görünür küçük değişiklik (örn. başlık metni) ile `v1.0.1` release.
3. Kurulu app aç → "yeni sürüm" diyalog çıkmalı.
4. Onayla → restart → değişiklik görünmeli.
5. Bilerek bozuk `v1.0.2` (örn. `main.py`'de `import yokmodul`) release.
6. Aç → bir an açılır gibi olur, rollback tetiklenir, eski sürüm açılır.
7. **Veya:** lokalde `src/main.py`'yi elle boz → exe aç → backup geri yüklenmeli.

### 7.3 Sürüm release süreci (geliştirici akışı)

1. Claude/Yavuz: `src/main.py` (veya hangi dosya) düzenlenir.
2. Lokal test yapılır.
3. `src/version.txt` artırılır (`1.0.0` → `1.0.1`).
4. `cd src && zip -r ../src.zip *.py version.txt && cd ..`
5. `sha256sum src.zip` → release notes'a yazılır.
6. `gh release create v1.0.1 src.zip --notes "..."` (release notes'a Türkçe değişiklik özeti + SHA256).
7. Bitti — şirket bilgisayarındaki app bir sonraki açılışta bulur.

---

## 8. Scope dışında bırakılanlar (YAGNI)

- **Delta updates** — toplam src ~50KB; tam zip indirmek yeterli, diff motoru gereksiz.
- **Otomatik onaysız update** — kullanıcı her zaman "Şimdi/Sonra" seçer; "zorunlu update" şu an YOK.
- **Birden fazla kullanıcı için telemetri / sürüm raporlama** — tek kullanıcı varken gereksiz.
- **Code signing certificate** — para gerektiriyor; mevcut AV durumunu görmeden almaya gerek yok.
- **Şirket içi mirror / offline update** — şirket internet açık olduğu sürece gereksiz.
- **Shotcrete ledger update'i** — bu spec sadece auto-update altyapısı; shotcrete planı ayrı spec.

---

## 9. Önkoşullar

- GitHub hesabı + public repo: `<github_user>/CubeLogReader` (bir kerelik).
- Geliştirici makinesinde `gh` CLI: `winget install --id GitHub.cli` (bir kerelik, sonra `gh auth login`).
- Şirket bilgisayarında: değişiklik yok; sadece internet erişimi (zaten var).

---

## 10. Implementation sırası (üst-düzey)

1. `src/updater.py` yaz + birim testleri.
2. `launcher.py` yaz + rollback testi (lokal, manuel bozma ile).
3. `src/main.py`'ye update check entegrasyonu + SettingsDialog butonu.
4. `CubeLogReader.spec` ve `build_exe.bat` güncelle, yeni klasör yapısıyla build dene.
5. `installer.iss` `[Files]` güncelle.
6. GitHub repo oluştur, ilk `v1.0.0` release at.
7. End-to-end test (yukarıdaki 7.2 senaryosu).
8. Şirket bilgisayarına yeni installer kur (sadece bir kerelik).

Detaylı plan **writing-plans** skill ile sonra çıkarılacak.
