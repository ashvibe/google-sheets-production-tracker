# 🏭 Arora Üretim Takip Sistemi

Google Sheets + Apps Script tabanlı üretim yönetim sistemi. Siparişleri, şase teslimatlarını, üretim emirlerini (ARP) ve bant çıktısını tek bir elektronik tabloda sidebar formları ve canlı dashboard ile takip eder.

> 🇬🇧 For English documentation: [README.md](./README.md)

---

## ✨ Özellikler

- **Sipariş Yönetimi** — Sipariş numarası başına çok kalemli sipariş, otomatik artan numara
- **Şase Takibi** — Tedarikçi bazlı takip ve teslimat logu (her sevkiyat ayrı satırda)
- **ARP (Üretim Emirleri)** — Modelleri bantlara atama, planlanan vs. gerçekleşen üretim takibi
- **Bant Takibi** — Bant başına günlük üretim girişi (A/B/C/D)
- **Dashboard** — KPI kartları, haftalık bant üretim tablosu (gün bazlı), sipariş ilerleme çubukları
- **Sidebar Formlar** — Ham hücre düzenleme yerine kullanıcı dostu formlar
- **Yedekleme & Arşivleme** — Tek tıkla Google Drive yedekleme; geçmişi kaybetmeden bant verisini temizleme
- **Şifreli Sıfırlama** — Şifre korumalı sistem sıfırlama (silmeden önce otomatik yedek)
- **Ayarlar Sayfası** — Koda dokunmadan model, tedarikçi ve bant ekleme

---

## 🗂️ Sayfa Yapısı

| Sayfa | Amaç | Birincil Kullanıcı |
|-------|------|--------------------|
| 📊 DASHBOARD | KPI özeti, haftalık bant tablosu, sipariş ilerlemesi | Yönetim |
| 📋 SİPARİŞLER | Tüm siparişler, durum ve şase bilgisi | Yönetim |
| 🔩 ŞASE TAKİP | Tedarikçi bazlı şase teslimat logu | Satın Alma |
| ⚙️ ARP | Bant atamaları ve üretim takibi | Üretim Sorumlusu |
| 📈 BANT TAKİP | Bant başına günlük üretim girişleri | Bant Operatörleri |
| ⚙️ AYARLAR | Model, tedarikçi ve bant listeleri | Admin |

---

## 🚀 Kurulum

### Gereksinimler
- Google hesabı
- Google Sheets (ücretsiz)

### Adımlar

1. Yeni bir Google Sheets dosyası açın
2. **Uzantılar → Apps Script** gidin
3. Varsayılan kodu silin
4. `tr/UretimTakip.gs` dosyasının içeriğini yapıştırın
5. **Kaydet** (💾) tıklayın
6. Fonksiyon listesinden `kurulumYap` seçin ve ▶️ çalıştırın
7. İzin istediğinde onaylayın
8. Sheets'e dönün — **🏭 Üretim Takip** menüsü belirecek

---

## 📋 Kullanım

### Sipariş Ekleme (Yönetim)
**Menü → Yeni Sipariş Ekle**
- *Yeni Sipariş* (otomatik numara) veya *Mevcut Siparişe Kalem Ekle* seçin
- Model seçin, adet girin, tarihleri ayarlayın
- Sipariş hem Siparişler hem Şase Takip sayfasına otomatik eklenir

### Şase Girişi (Satın Alma)
**Menü → Şase Girişi Yap**
- Model, tedarikçi seçin; gelen adedi ve tarihleri girin
- Özet satırı otomatik güncellenir (gelen toplam, kalan, durum)
- Her sevkiyat Giriş Logu'na ayrı satır olarak eklenir

### ARP Oluşturma (Üretim Sorumlusu)
**Menü → Yeni ARP Oluştur**
- Sipariş, model, bant, planlanan adeti seçin
- ARP numarası otomatik atanır (ARP-0001, ARP-0002...)
- Sipariş durumu "Üretimde" olarak güncellenir

### Günlük Bant Girişi (Operatörler)
**Menü → Bant Girişi**
- Tarih, hat, model ve üretilen adeti girin
- İsteğe bağlı: ARP numarasıyla ilişkilendirin

### Dashboard Yenileme
**Menü → Dashboard Yenile**
- Tüm KPI kartları güncellenir
- Haftalık üretim tablosu yeniden oluşturulur
- Sipariş ilerleme çubukları güncellenir

---

## 💾 Yedekleme & Arşivleme

| İşlem | Menü Seçeneği | Ne Yapar |
|-------|--------------|----------|
| Manuel yedek | 📥 Yedek Al | Tüm dosyayı zaman damgalı olarak Drive'a kopyalar |
| Bant arşivleme | 🧹 Bant Takibi Arşivle | Önce yedek alır, sonra bant satırlarını temizler |
| Sistem sıfırlama | 🔧 Sistemi Kur / Sıfırla | Şifre korumalı; silmeden önce otomatik yedek alır |

**Varsayılan admin şifresi:** Kodda `"ARORA00"` yazan yeri değiştirin.

---

## ⚙️ Yapılandırma

### Model Ekleme
1. **Menü → Ayarlar**
2. **A sütununun** altına yeni model adlarını yazın
3. **✅ Bitti** tıklayın — tüm dropdownlar otomatik güncellenir

### Tedarikçi Ekleme
1. **Menü → Ayarlar**
2. **C sütununa** tedarikçi adlarını ekleyin

### Admin Şifresini Değiştirme
Kodda `"ARORA00"` yazan yeri kendi şifrenizle değiştirin.

---

## 📁 Repo Yapısı

```
arora-production-tracker/
├── en/
│   └── ProductionTracker.gs     # İngilizce versiyon
├── tr/
│   └── UretimTakip.gs           # Türkçe versiyon
├── docs/
│   └── screenshots/             # Ekran görüntüleri
├── README.md                    # İngilizce dokümantasyon
└── README.tr.md                 # Türkçe dokümantasyon (bu dosya)
```

---

## 🔑 ARP Numaralandırması

Üretim emirleri `ARP-0001`, `ARP-0002` formatını izler.

*ARP = Arora Production* — sektörde kullanılan TSM (Temsa Service Management) isimlendirme geleneğinden ilham alınmıştır.

---

## 📄 Lisans

MIT Lisansı — özgürce kullanabilir, değiştirebilir ve dağıtabilirsiniz.

---

*Google Apps Script ile geliştirilmiştir. Harici bağımlılık yoktur.*
