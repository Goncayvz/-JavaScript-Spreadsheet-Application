# 📊 Fonksiyonel Spreadsheet Uygulaması

<div align="center">

![Spreadsheet Banner](https://img.shields.io/badge/Spreadsheet-Application-blue?style=for-the-badge)
![JavaScript](https://img.shields.io/badge/JavaScript-ES6+-yellow?style=for-the-badge&logo=javascript)
![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=for-the-badge&logo=html5&logoColor=white)
![CSS3](https://img.shields.io/badge/CSS3-1572B6?style=for-the-badge&logo=css3&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

**Saf JavaScript ile geliştirilmiş, fonksiyonel programlama yaklaşımlı modern bir spreadsheet uygulaması**

[Canlı Demo](#-canlı-demo) • [Özellikler](#-özellikler) • [Kurulum](#-kurulum) • [Kullanım](#-kullanım)
#Demo 
 https://goncayvz.github.io/-JavaScript-Spreadsheet-Application
</div>

---

## 🎯 Hakkında

Bu proje, **[freeCodeCamp](https://www.freecodecamp.org/)** JavaScript Algorithms and Data Structures sertifikasyon programındaki Spreadsheet projesini baz alarak geliştirilmiş, **tamamen yeniden tasarlanmış ve genişletilmiş** modern bir web uygulamasıdır.

### 🌟 Neden Bu Proje?

- ✅ **Sıfır framework/library** - Saf JavaScript ile yazılmış
- ✅ **Fonksiyonel programlama** - Modern JavaScript best practices
- ✅ **Excel-benzeri deneyim** - Tanıdık arayüz ve özellikler
- ✅ **Eğitim amaçlı** - Kodlar temiz ve yorumlu
- ✅ **Açık kaynak** - Özgürce kullanın ve geliştirin

---

## ✨ Özellikler

### 📊 Temel Özellikler

- **990 Hücreli Grid** (A-J sütunları, 1-99 satırları)
- **Excel Formül Motoru** - 20+ yerleşik fonksiyon
- **Hücre Formatlama** - Formüller, sayılar, metin
- **Kopyala/Yapıştır** - Tam hücre kopyalama desteği
- **Geri Al/İleri Al** - Sınırsız undo/redo
- **CSV Dışa Aktarma** - Verilerinizi indirin

### 🎨 Gelişmiş Özellikler

#### 📈 Grafik Sistemi
- **3 Grafik Tipi**: Çubuk, Çizgi, Pasta
- **Akıllı Veri Analizi** - Otomatik grafik önerisi
- **PNG Export** - Grafikleri resim olarak indirin
- **Canvas-based** - Yüksek kaliteli render

#### 🤖 AI-Powered Formül Yardımı
- **Akıllı Formül Asistanı** - F1 ile hızlı erişim
- **40+ Formül Dokümantasyonu** - Detaylı açıklamalar ve örnekler
- **Kategorize Yardım** - Matematik, İstatistik, Metin, Tarih
- **Arama Özelliği** - İstediğiniz formülü bulun

#### 💬 Basit Chatbot
- **Spreadsheet Asistanı** - Formül yardımı
- **Önceden tanımlı cevaplar** - Hızlı yanıtlar
- **Hızlı soru önerileri** - Tek tıkla sorular

#### 🎨 Tema Sistemi
- **Karanlık/Aydınlık Mod** - Göz yorgunluğunu azaltın
- **Otomatik Algılama** - Sistem temasını takip eder
- **Tercih Kaydetme** - Seçiminiz saklanır

#### ⚡ Performans & UX
- **Hata Yönetimi** - 8 farklı hata tipi desteği
- **Gerçek Zamanlı Hesaplama** - Anlık formül güncellemesi
- **Performans İzleme** - Detaylı istatistikler
- **Klavye Kısayolları** - Hızlı navigasyon

---

## 🚀 Kurulum

### Gereksinimler

- Modern bir web tarayıcısı (Chrome, Firefox, Safari, Edge)
- Yerel bir HTTP sunucusu (opsiyonel)

### Hızlı Başlangıç

```bash
# 1. Projeyi klonlayın
git clone https://github.com/kullaniciadi/fonksiyonel-spreadsheet.git
cd fonksiyonel-spreadsheet

# 2. Tarayıcıda açın
# Basit yöntem
open index.html

# VEYA yerel sunucu kullanın
python -m http.server 8000
# http://localhost:8000
```

---

## 📖 Kullanım

### Temel Kullanım

#### 1️⃣ Hücre Seçimi
- **Tek tıklama** - Tek hücre seçimi
- **Sürükleme** - Çoklu hücre seçimi
- **Ctrl + Click** - Çoklu seçim
- **Ok tuşları** - Klavye ile gezinme

#### 2️⃣ Veri Girişi
```javascript
// Sayı
42

// Metin
Merhaba Dünya

// Formül
=SUM(A1:A10)
```

#### 3️⃣ Formül Kullanımı
```javascript
// Toplama
=SUM(A1:A5)

// Ortalama
=AVERAGE(B1:B10)

// Koşullu
=IF(C1>100, "Yüksek", "Düşük")
```

### Klavye Kısayolları

| Kısayol | Açıklama |
|---------|----------|
| `Enter` | Hücreyi düzenle |
| `ESC` | İptal et |
| `Tab` | Sağ hücre |
| `Ctrl + C` | Kopyala |
| `Ctrl + V` | Yapıştır |
| `Ctrl + Z` | Geri al |
| `Ctrl + Y` | İleri al |
| `Ctrl + S` | CSV kaydet |
| `F1` | Formül yardımı |
| `F9` | Test |

---

## 📐 Desteklenen Formüller

### 🔢 Matematik

| Fonksiyon | Açıklama | Örnek |
|-----------|----------|-------|
| `SUM` | Toplama | `=SUM(A1:A10)` |
| `AVERAGE` | Ortalama | `=AVERAGE(B1:B5)` |
| `MAX` | En büyük | `=MAX(C1:C10)` |
| `MIN` | En küçük | `=MIN(D1:D10)` |
| `POWER` | Üs | `=POWER(2, 3)` |
| `SQRT` | Karekök | `=SQRT(16)` |
| `ROUND` | Yuvarlama | `=ROUND(3.14, 2)` |
| `ABS` | Mutlak değer | `=ABS(-5)` |

### 📊 İstatistik

| Fonksiyon | Açıklama | Örnek |
|-----------|----------|-------|
| `MEDIAN` | Medyan | `=MEDIAN(A1:A10)` |
| `STDEV` | Std. sapma | `=STDEV(B1:B10)` |
| `COUNT` | Sayma | `=COUNT(C1:C10)` |

### 🎯 Mantık

| Fonksiyon | Açıklama | Örnek |
|-----------|----------|-------|
| `IF` | Koşul | `=IF(A1>10, "Büyük", "Küçük")` |
| `AND` | Ve | `=AND(A1>0, A1<100)` |
| `OR` | Veya | `=OR(B1="A", B1="B")` |

### 📝 Metin

| Fonksiyon | Açıklama | Örnek |
|-----------|----------|-------|
| `CONCAT` | Birleştir | `=CONCAT(A1, B1)` |
| `LEN` | Uzunluk | `=LEN(A1)` |
| `UPPER` | Büyük harf | `=UPPER("a")` |

---

## 🛠️ Teknolojiler

- **HTML5** - Semantik yapı
- **CSS3** - Modern styling (Grid, Flexbox)
- **Vanilla JavaScript (ES6+)** - Sıfır framework
- **Canvas API** - Grafik çizimi
- **LocalStorage API** - Veri saklama
- **Font Awesome 6.4.0** - İkonlar

---

## 📁 Proje Yapısı

```
fonksiyonel-spreadsheet/
│
├── index.html           # Ana HTML
├── styles.css           # Ana CSS
├── script.js            # Ana JavaScript
├── errorHandler.js      # Hata yönetimi
│
└── README.md           # Bu dosya
```

---

## 💻 Geliştirme

### Debug Modu

```javascript
// Console'da
window.ENV.DEBUG_MODE = true;

// Performans göster
showPerformance();

// Test çalıştır
runExcelTest();
```

---

## 🤝 Katkıda Bulunma

1. Fork edin
2. Feature branch oluşturun (`git checkout -b feature/amazing`)
3. Commit edin (`git commit -m 'feat: Add feature'`)
4. Push edin (`git push origin feature/amazing`)
5. Pull Request açın

### Commit Kuralları

```bash
feat: Yeni özellik
fix: Bug düzeltme
docs: Dokümantasyon
style: Kod formatı
refactor: Kod iyileştirme
test: Test ekleme
perf: Performans
```
---

## 📊 Proje İstatistikleri

```
📝 Toplam Satır     : ~5,000
🎨 CSS              : ~1,800
💻 JavaScript       : ~3,000
📁 Dosya            : 4
📦 Bağımlılık       : 0
```

---


## 🙏 Teşekkürler

- **[freeCodeCamp](https://www.freecodecamp.org/)** - Orijinal proje
- **[MDN Web Docs](https://developer.mozilla.org/)** - Dokümantasyon
- **[Font Awesome](https://fontawesome.com/)** - İkonlar

---

## 📞 İletişim

**Proje Sahibi:** [Adınız]

- 📧 Email: email@example.com
- 🐙 GitHub: [@kullaniciadi](https://github.com/kullaniciadi)
- 💼 LinkedIn: [linkedin.com/in/kullaniciadi](https://linkedin.com/in/kullaniciadi)

**Proje Linki:** [github.com/kullaniciadi/fonksiyonel-spreadsheet](https://github.com/kullaniciadi/fonksiyonel-spreadsheet)

---

## ⭐ Yıldızlayın!

Projeyi beğendiyseniz ⭐ vermeyi unutmayın!

<div align="center">

**Made with ❤️ and ☕**

![Footer](https://img.shields.io/badge/Thanks%20for-Visiting-blue?style=for-the-badge)

</div>
