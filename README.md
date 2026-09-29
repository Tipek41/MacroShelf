<div align="right">

[English](#english) · [Türkçe](#türkçe)

</div>

# MacroShelf

<a id="english"></a>

## English

A growing collection of practical Excel VBA macros to automate everyday tasks.

### Available macros

#### 1. Merge PDFs in a folder

Select a folder and merge its top-level PDF files into one PDF, sorted by filename. Source PDFs stay unchanged. The macro skips its previous output and reports files it could not read.

- [Download the VBA module](pdf-merge.bas)
- **Requirements:** Excel for Windows and the Adobe Acrobat desktop app. Acrobat Reader alone does not support the document-editing automation used by this macro.
- Files are processed on your computer; PDFs are not uploaded to a web service.

#### 2. Clear formatting outside the data

Clear formatting beyond the last value or formula on the active worksheet. Formulas that currently return blank text are included in the data boundary. Cell values and formulas stay intact; the macro does not delete or shift rows and columns.

- [Download the VBA module](clear-unused-formatting.bas)
- **Important:** Make a copy of the workbook before running it. The macro changes formatting, which Excel may not be able to undo.
- If Excel still reports an inflated used range, save and reopen the workbook.

### Install and run

1. In Excel, press **Alt+F11** to open the VBA editor.
2. Choose **File → Import File…** and select the `.bas` module for the macro you want.
3. Press **Alt+F8**, select the macro, and run it. PDF merging asks you to choose a folder; format cleanup works on the active worksheet.

### Website

The bilingual MacroShelf page: [tipek41.github.io/MacroShelf](https://tipek41.github.io/MacroShelf/).

---

<a id="türkçe"></a>

## Türkçe

Günlük işleri otomatikleştiren kullanışlı Excel VBA makrolarının büyüyen koleksiyonu.

### Kullanılabilir makrolar

#### 1. Klasördeki PDF’leri birleştir

Bir klasör seçin; alt klasörlere girmeden içindeki PDF’ler dosya adına göre sıralanıp tek PDF’te birleştirilsin. Kaynak dosyalar değiştirilmez. Makro önceki çıktı dosyasını tekrar eklemez ve okuyamadığı dosyaları bildirir.

- [VBA modülünü indir](pdf-merge.bas)
- **Gereksinimler:** Windows için Excel ve Adobe Acrobat masaüstü uygulaması. Acrobat Reader, bu makronun kullandığı PDF düzenleme otomasyonunu desteklemez.
- Dosyalar bilgisayarınızda işlenir; PDF’ler bir web servisine yüklenmez.

#### 2. Veri dışındaki biçimleri temizle

Aktif çalışma sayfasındaki son değer veya formül hücresinin dışındaki biçimleri temizler. Şu anda boş metin döndüren formüller de veri sınırına dahil edilir. Hücre değerleri ve formüller korunur; makro satır ya da sütun silmez, hücreleri kaydırmaz.

- [VBA modülünü indir](clear-unused-formatting.bas)
- **Önemli:** Çalıştırmadan önce çalışma kitabının bir kopyasını alın. Biçim değişiklikleri Excel’de geri alınamayabilir.
- Excel kullanılan alanı hâlâ geniş gösteriyorsa dosyayı kaydedip kapatın ve yeniden açın.

### Kurulum ve çalıştırma

1. Excel’de **Alt+F11** tuşlarına basarak VBA düzenleyicisini açın.
2. **File → Import File…** seçeneğinden kullanmak istediğiniz makronun `.bas` dosyasını içe aktarın.
3. **Alt+F8** tuşlarına basıp makroyu seçerek çalıştırın. PDF birleştirme klasör seçtirir; biçim temizleme aktif çalışma sayfasında çalışır.

### Web sitesi

İki dilli MacroShelf sayfası: [tipek41.github.io/MacroShelf](https://tipek41.github.io/MacroShelf/)

---

[↑ Back to English](#english) · [↑ Türkçeye dön](#türkçe)
