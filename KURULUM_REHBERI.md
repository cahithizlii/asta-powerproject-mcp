# Asta Powerproject MCP Server - Kurulum Rehberi

## Bu Nedir?

Bu, Claude Desktop uygulamasinin Asta Powerproject ile calismasi icin ozel olarak gelistirilmis bir MCP (Model Context Protocol) sunucusudur.

Kuruldugunda, Claude su islemleri yapabilir:

### Dosya Islemleri (MPXJ ile):
- Proje dosyalarini (.pp, .mpp, .xml) okuma ve analiz etme
- Gorevleri listeleme, ekleme, guncelleme, silme
- Kritik yol analizi
- Float (kayma payi) analizi
- Kaynak ve maliyet bilgilerini goruntuleme
- Gorev baglantilarini (links) goruntuleme
- Projeyi XML formatinda kaydetme

### GUI Otomasyon (Asta penceresini kontrol):
- Asta'nin acik olup olmadigini kontrol etme
- Reschedule (F9) calistirma
- Kaydetme (Ctrl+S)
- Ekran goruntusu alma
- Hucrelere veri girisi
- Filtre uygulama
- PDF/resim olarak cikti alma
- Baseline alma
- Gorev baglama

---

## Kurulum Adimlari (Adim Adim)

### Adim 1: Python Kurulumu

1. Web tarayicida su adrese gidin: **https://www.python.org/downloads/**
2. "Download Python 3.12" (veya en yeni surum) butonuna tiklayin
3. Indirilen dosyayi calistirin
4. **ONEMLI:** Kurulum ekraninda **"Add Python to PATH"** kutucugunu isaretleyin
5. "Install Now" tiklayin
6. Kurulum bitene kadar bekleyin

**Kontrol:** Komut Istemi (Command Prompt) acin ve yazin:
```
python --version
```
"Python 3.12.x" gibi bir cikti gormelisiniz.

### Adim 2: Dosyalari Yerlestime

1. `asta-powerproject-mcp` klasorunu bilgisayarinizda uygun bir yere kopyalayin
   - Onerim: `C:\Users\GPX PRO\asta-powerproject-mcp\`
   - Ya da: `C:\Tools\asta-powerproject-mcp\`

2. Klasorun icinde su dosyalar olmali:
   - `asta_mcp_server.py` (ana sunucu dosyasi)
   - `requirements.txt` (gerekli paketler listesi)
   - `install.bat` (otomatik kurulum scripti)
   - `claude_desktop_config_example.json` (ornek config)

### Adim 3: Paketleri Yukleme

**Yontem A - Otomatik (Kolay):**
1. `asta-powerproject-mcp` klasorune gidin
2. `install.bat` dosyasina cift tiklayin
3. Kurulum otomatik olarak yapilacak
4. "Installation Complete!" mesajini gorun

**Yontem B - Manuel:**
1. Komut Istemi'ni yonetici olarak acin (CMD)
2. Klasore gidin:
   ```
   cd C:\Users\GPX PRO\asta-powerproject-mcp
   ```
3. Paketleri yukleyin:
   ```
   pip install -r requirements.txt
   ```
4. Kurulumu dogrulayin:
   ```
   python -c "from mcp.server.fastmcp import FastMCP; print('Basarili!')"
   ```

### Adim 4: Claude Desktop Yapilandirmasi

1. Claude Desktop uygulamasini acin
2. Sol alt kosedeki **ayarlar (disli) ikonuna** tiklayin
3. **"Developer"** sekmesine gidin
4. **"Edit Config"** butonuna tiklayin
5. Acilan dosyada asagidaki icerigi ekleyin:

```json
{
  "mcpServers": {
    "asta_powerproject_mcp": {
      "command": "python",
      "args": ["C:\\Users\\GPX PRO\\asta-powerproject-mcp\\asta_mcp_server.py"]
    }
  }
}
```

**ONEMLI:** `args` icindeki yolu kendi dosya yolunuzla degistirin!
Ters slash (\) yerine cift ters slash (\\) kullanin.

6. Dosyayi kaydedin (Ctrl+S)
7. Claude Desktop'u **tamamen kapatin** ve yeniden acin

### Adim 5: Dogrulama

1. Claude Desktop'u acin
2. Yeni bir konusma baslatın
3. Sol alt kosede arac simgelerine (tools) bakin
4. "asta_powerproject_mcp" gorulmeli
5. Test icin Claude'a sorun: "Asta MCP server'i test et"

---

## Kullanim Ornekleri

### Proje Analizi:
> "C:/Users/GPX PRO/Downloads/proje.pp dosyasini analiz et"

### Gorev Listesi:
> "Projedeki tum gorevleri listele"

### Kritik Yol:
> "Kritik yol analizini goster"

### Gorev Ekleme:
> "Projeye 'Beton Dokum' adinda 5 gunluk yeni gorev ekle"

### Ilerleme Guncelleme:
> "Task 15'in ilerleme yuzdesi %60 olarak guncelle"

### Reschedule:
> "Asta'da projeyi reschedule yap (F9)"

### Ekran Goruntusu:
> "Asta'nin ekran goruntusu al"

---

## Sorun Giderme

### "Python bulunamadi" hatasi:
- Python'u yeniden yukleyin
- "Add to PATH" secenegini isaretleyin
- Bilgisayari yeniden baslatin

### "mcp modulu bulunamadi" hatasi:
```
pip install mcp
```

### "mpxj yuklenemiyor" hatasi:
- Java JDK 11+ gerekebilir: https://adoptium.net/
- Java yuklendikten sonra: `pip install mpxj`

### "pyautogui/pywinauto yuklenemiyor" hatasi:
- Bu paketler sadece GUI otomasyonu icindir
- Dosya islemleri bunlar olmadan da calisir
- Yuklemek icin: `pip install pyautogui pywinauto`

### Claude Desktop'ta arac gorunmuyor:
1. Config dosyasinin dogru oldugunu kontrol edin
2. Yol icindeki ters slashlari (\\) kontrol edin
3. Claude Desktop'u tamamen kapatip yeniden acin
4. Log dosyasini kontrol edin: `C:\Users\GPX PRO\asta_mcp.log`

### "Permission denied" hatasi:
- Komut Istemi'ni yonetici olarak calistirin
- Veya: `pip install --user -r requirements.txt`

---

## Dosya Yapisi

```
asta-powerproject-mcp/
├── asta_mcp_server.py              # Ana MCP sunucu dosyasi
├── requirements.txt                 # Gerekli Python paketleri
├── install.bat                      # Windows otomatik kurulum
├── claude_desktop_config_example.json  # Ornek config dosyasi
└── KURULUM_REHBERI.md              # Bu dosya
```

---

## Desteklenen Dosya Formatlari

| Format | Uzanti | Aciklama |
|--------|--------|----------|
| Asta Powerproject | .pp | Ana proje dosyasi |
| Microsoft Project | .mpp | MS Project dosyasi |
| MS Project XML | .xml, .mspdi | XML formati |
| Primavera XER | .xer | Oracle Primavera |
| Primavera XML | .pmxml | Primavera P6 XML |

---

## Araclarin Tam Listesi

### Dosya Araclari:
| Arac | Aciklama |
|------|----------|
| asta_analyze_project | Proje dosyasini analiz et |
| asta_list_tasks | Tum gorevleri listele |
| asta_get_task | Tek gorev detayi |
| asta_add_task | Yeni gorev ekle |
| asta_update_task | Gorevi guncelle |
| asta_delete_task | Gorevi sil |
| asta_get_critical_path | Kritik yol analizi |
| asta_list_resources | Kaynaklari listele |
| asta_get_resource_assignments | Kaynak atamalari |
| asta_get_calendars | Takvimleri goster |
| asta_float_analysis | Float (kayma) analizi |
| asta_save_project | Projeyi kaydet |

### GUI Araclari:
| Arac | Aciklama |
|------|----------|
| asta_gui_check_status | Asta durumunu kontrol et |
| asta_gui_bring_to_front | Asta'yi one getir |
| asta_gui_send_shortcut | Klavye kisayolu gonder |
| asta_gui_reschedule | Reschedule (F9) |
| asta_gui_save | Kaydet (Ctrl+S) |
| asta_gui_undo | Geri al (Ctrl+Z) |
| asta_gui_click | Koordinata tikla |
| asta_gui_type_text | Metin yaz |
| asta_gui_screenshot | Ekran goruntusu al |
| asta_gui_open_file | Dosya ac |
| asta_gui_new_project | Yeni proje olustur |
| asta_gui_take_baseline | Baseline al |
| asta_gui_insert_row | Satir ekle |
| asta_gui_delete_selected | Secili oge sil |
| asta_gui_link_tasks | Gorevleri bagla |
| asta_gui_apply_filter | Filtre uygula |
| asta_gui_change_table | Tablo gorunumu degistir |
| asta_gui_print_export | Yazdir/PDF cikti |
| asta_gui_zoom | Yakınlastir/uzaklastir |
| asta_gui_summarize_tasks | Ozet gorev olustur |
| asta_gui_indent_task | Iceri/disari tasi |
| asta_help | Yardim rehberi |

---

**Versiyon:** 1.0.0
**Tarih:** Subat 2026
**Hazırlayan:** Claude AI
