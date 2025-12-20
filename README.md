# 📊 JavaScript Spreadsheet Application

**Gelişmiş Web Tabanlı Spreadsheet (Elektronik Tablo) Uygulaması**

Bu proje, **freeCodeCamp Spreadsheet projesi** temel alınarak geliştirilmiş; tamamen **Vanilla JavaScript, HTML ve CSS** kullanılarak oluşturulmuş modern, hızlı ve fonksiyonel bir web tabanlı spreadsheet uygulamasıdır.

Excel benzeri formül hesaplama motoru, gelişmiş hücre yönetimi ve kullanıcı dostu arayüz özelliklerini bir araya getirir.

---

## 🚀 Özellikler

### 📊 Gelişmiş Hesaplama Motoru

* Excel uyumlu formüller
  `SUM`, `AVERAGE`, `MAX`, `MIN`, `COUNT`, `MEDIAN`
* Dinamik hücre referansları
  `A1`, `B2`, `A1:A10` gibi aralık desteği
* Gerçek zamanlı hesaplama
* Gelişmiş hata yönetimi:

  * Sıfıra bölme
  * Syntax hataları
  * Geçersiz referanslar
  * Döngüsel (sonsuz) referanslar

---

### 🎨 Modern & Responsive Arayüz

* 🌙 Koyu / ☀️ Açık tema desteği
* Çoklu hücre seçimi

  * `Ctrl + Click`
  * Sürükle & bırak
* Formül çubuğu
* Durum çubuğu (seçili hücre bilgileri)
* Tam klavye navigasyonu:

  * Ok tuşları
  * Enter / Tab
  * F2 ile düzenleme

---

### 🔧 Profesyonel Araçlar

* Kopyala / Yapıştır (`Ctrl + C`, `Ctrl + V`)
* Geri Al / İleri Al (`Ctrl + Z`, `Ctrl + Y`)
* CSV dışa aktarma
* Demo veri yükleme
* Performans ve kullanım istatistikleri

---

## 🛠️ Kurulum

Depoyu klonlayın:

```bash
git clone https://github.com/Goncayvz/-JavaScript-Spreadsheet-Application.git
```

Ardından `index.html` dosyasını bir tarayıcıda açmanız yeterlidir.
Herhangi bir ek bağımlılık veya kurulum gerektirmez.

---

## 🧩 Temel Kullanım

* **Hücre Seçimi:** Tıklayarak veya ok tuşları ile
* **Veri Girişi:** Seçili hücreye doğrudan yazın
* **Formül Kullanımı:** `=` ile başlayın
  Örnek: `=SUM(A1:A5)`
* **Düzenleme Modu:** Çift tıklayın veya `F2`
* **Onaylama:** `Enter`

---

## ⌨️ Klavye Kısayolları

| Kısayol     | Açıklama                  |
| ----------- | ------------------------- |
| Ctrl + C    | Kopyala                   |
| Ctrl + V    | Yapıştır                  |
| Ctrl + Z    | Geri al                   |
| Ctrl + Y    | İleri al                  |
| Ctrl + S    | CSV olarak indir          |
| F2          | Hücreyi düzenle           |
| F9          | Excel uyumluluk testi     |
| ESC         | Düzenlemeyi iptal et      |
| Tab         | Sağdaki hücre             |
| Shift + Tab | Soldaki hücre             |
| Enter       | Kaydet ve alt hücreye geç |

---

## 📐 Desteklenen Formüller

| Fonksiyon | Açıklama             | Örnek             |
| --------- | -------------------- | ----------------- |
| SUM       | Toplama              | `=SUM(A1:A10)`    |
| AVERAGE   | Ortalama             | `=AVERAGE(B1:B5)` |
| MAX       | Maksimum             | `=MAX(C1:C20)`    |
| MIN       | Minimum              | `=MIN(D1:D15)`    |
| COUNT     | Sayısal hücre sayısı | `=COUNT(E1:E100)` |
| MEDIAN    | Medyan               | `=MEDIAN(F1:F10)` |

---

## ⚠️ Hata Türleri

| Hata Kodu             | Açıklama                 |
| --------------------- | ------------------------ |
| `#SYNTAX`             | Formül sözdizimi hatası  |
| `#REFERENCE`          | Geçersiz hücre referansı |
| `#DIV_ZERO`           | Sıfıra bölme             |
| `#CALC_TIMEOUT`       | Hesaplama zaman aşımı    |
| `#CALC_INFINITE_LOOP` | Döngüsel referans        |

---

## 🧠 Teknik Detaylar

### Mimari

* Vanilla JavaScript (harici kütüphane yok)
* Fonksiyonel programlama yaklaşımı
* Modüler dosya yapısı
* Event-driven mimari

### Performans

* Hesaplama önbelleği (cache)
* Optimize DOM güncellemeleri
* Bellek sızıntısı önleme
* Debounced input işleme

---

## 🌐 Tarayıcı Uyumluluğu

* Chrome 90+
* Firefox 88+
* Edge 90+
* Safari 14+
* Mobil & Tablet uyumlu
* ARIA destekli erişilebilirlik

---

## 📊 Demo Veri Seti

Uygulama, hızlı test ve öğrenme için hazır demo verileri içerir:

* Satış verileri
* Toplam & ortalama hesaplamaları
* Tüm hata türlerine örnekler
* Formül kullanım senaryoları

**Demo verilerini yüklemek için:**
👉 *“Demo Veriler”* butonuna tıklayın.

---

## 📄 Lisans

Bu proje eğitim ve geliştirme amaçlıdır.
Dilediğiniz gibi kullanabilir, geliştirebilir ve paylaşabilirsiniz.

---

💡 *Her türlü geri bildirim ve katkıya açıktır.*
