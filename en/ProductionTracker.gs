// ============================================================
// ARORA PRODUCTION TRACKER v2
// Setup: Run kurulumYap() once to initialize.
// GitHub: https://github.com/yourusername/arora-production-tracker
// ============================================================

// ── SABITLER ─────────────────────────────────────────────────
const RENKLER = {
  KOYU:    "#1a1a2e",
  BEYAZ:   "#ffffff",
  YESIL:   "#d4edda",
  SARI:    "#fff3cd",
  KIRMIZI: "#f8d7da",
  MAVI:    "#d1ecf1",
  GRI:     "#f0f0f0",
  TURUNCU: "#ffeaa7",
  BASLIK:  "#16213e",
  METIN:   "#e0e0e0",
};

const SIPARIS_STATUS  = ["Waiting","In Production","Completed","Cancelled"];
const SASE_STATUS     = ["Not Ordered","Ordered","Partial","Fully Received"];
const IE_STATUS       = ["Planned","In Progress","Paused","Completed"];
const LINES        = ["A BANT","B BANT","C BANT","D BANT"];
const ONCELIK        = ["1-Critical","2-High","3-Normal","4-Low"];

// ── KURULUM ──────────────────────────────────────────────────


// ── YEDEKLEME ────────────────────────────────────────────────────────────────
function yedekAl() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const onay = ui.alert(
    "📥 YEDEK AL",
    "Tüm veriler Google Drive'ınıza yeni bir dosya olarak kaydedilecek.\n\nDevam edilsin mi?",
    ui.ButtonSet.YES_NO
  );
  if (onay !== ui.Button.YES) return;

  try {
    const tarih = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy_HH.mm");
    const dosyaAdi = `UretimYedek_${tarih}`;

    // Mevcut dosyayı kopyala
    const kaynak = DriveApp.getFileById(ss.getId());
    const kopya = kaynak.makeCopy(dosyaAdi);

    // Kopyayı sadece veri olarak tut (script olmadan)
    const yedekUrl = kopya.getUrl();

    logYaz("YEDEKLeme", "YEDEK ALINDI", dosyaAdi);
    ui.alert(
      "✅ Yedek Alındı!",
      `Dosya adı: ${dosyaAdi}\n\nGoogle Drive'ınızın ana klasörüne kaydedildi.\n\nURL: ${yedekUrl}`,
      ui.ButtonSet.OK
    );
  } catch(e) {
    ui.alert("❌ Hata", "Yedek alınamadı: " + e.message, ui.ButtonSet.OK);
  }
}

function bantArsivle() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bant = ss.getSheetByName("📈 LINE TRACKING");

  // Kaç satır var?
  const veri = bant.getRange("A4:H500").getValues().filter(r => r[0]);
  if (veri.length === 0) {
    ui.alert("Bant takip sayfasında silinecek veri yok.");
    return;
  }

  const onay = ui.alert(
    "🧹 BANT TAKİP ARŞİVLE",
    `Bant takipte ${veri.length} satır veri var.\n\n` +
    "1. Önce otomatik yedek alınacak\n" +
    "2. Sonra bant takip sayfası temizlenecek\n\n" +
    "Devam edilsin mi?",
    ui.ButtonSet.YES_NO
  );
  if (onay !== ui.Button.YES) return;

  // Önce yedek al
  try {
    const tarih = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy_HH.mm");
    const dosyaAdi = `UretimYedek_BantTemizleme_${tarih}`;
    const kaynak = DriveApp.getFileById(ss.getId());
    kaynak.makeCopy(dosyaAdi);

    // Bant takibi temizle (başlıkları koru)
    bant.getRange("A4:H500").clearContent();
    SpreadsheetApp.flush();

    logYaz("YEDEKLeme", "BANT ARŞİVLENDİ", `${veri.length} satır silindi, yedek: ${dosyaAdi}`);
    ui.alert(
      "✅ Tamamlandı!",
      `${veri.length} satır arşivlendi ve temizlendi.\n\nYedek dosya: ${dosyaAdi}\nGoogle Drive ana klasörünüzde.`,
      ui.ButtonSet.OK
    );
  } catch(e) {
    ui.alert("❌ Hata", "İşlem başarısız: " + e.message, ui.ButtonSet.OK);
  }
}

function kurulumKorumal() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt(
    "🔐 ADMIN ACCESS REQUIRED",
    "This will DELETE ALL DATA and reset the system!\n\nEnter password to continue:",
    ui.ButtonSet.OK_CANCEL
  );
  if (r.getSelectedButton() !== ui.Button.OK) return;
  if (r.getResponseText().trim() !== "YOUR_ADMIN_PASSWORD") {
    ui.alert("❌ Wrong password! Operation cancelled.");
    return;
  }
  const onay = ui.alert(
    "⚠️ FINAL WARNING",
    "All data will be deleted!\n\nSıfırlamadan önce otomatik yedek alınacak.\nEmin misiniz?",
    ui.ButtonSet.YES_NO
  );
  if (onay !== ui.Button.YES) return;

  // Önce otomatik yedek al
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tarih = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy_HH.mm");
    const dosyaAdi = `UretimYedek_SIFIRLAMA_ONCESI_${tarih}`;
    DriveApp.getFileById(ss.getId()).makeCopy(dosyaAdi);
    ui.alert("✅ Yedek alındı: " + dosyaAdi + "\n\nŞimdi sistem sıfırlanıyor...");
  } catch(e) {
    const devam = ui.alert("⚠️ Backup failed!\n" + e.message + "\n\nDo you still want to reset?", ui.ButtonSet.YES_NO);
    if (devam !== ui.Button.YES) return;
  }
  kurulumYap();
}

function kurulumYap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _sayfalarOlustur(ss);
  SpreadsheetApp.flush();
  _kurSiparisler(ss);
  _kurSaseTakip(ss);
  _kurIsEmirleri(ss);
  _kurBantTakip(ss);
  _kurDashboard(ss);
  _kurAyarlar(ss);
  SpreadsheetApp.flush();

  // Sayfa sırası
  const sira = ["📊 DASHBOARD","📋 ORDERS","🔩 CHASSIS TRACKING","⚙️ PRODUCTION ORDERS","📈 LINE TRACKING","⚙️ SETTINGS"];
  sira.forEach((ad, i) => {
    const s = ss.getSheetByName(ad);
    if (s) { ss.setActiveSheet(s); ss.moveActiveSheet(i + 1); }
  });
  ss.getSheetByName("⚙️ SETTINGS").hideSheet();
  ss.setActiveSheet(ss.getSheetByName("📊 DASHBOARD"));
  SpreadsheetApp.getUi().alert("✅ System installed! Use the menu to perform actions.");
}

function _sayfalarOlustur(ss) {
  const hedef = ["📊 DASHBOARD","📋 ORDERS","🔩 CHASSIS TRACKING","⚙️ PRODUCTION ORDERS","📈 LINE TRACKING","⚙️ SETTINGS"];
  hedef.forEach(ad => {
    let s = ss.getSheetByName(ad);
    if (!s) s = ss.insertSheet(ad);
    else s.clear();
  });
  // Eski sayfaları sil
  ss.getSheets().forEach(s => {
    if (!hedef.includes(s.getName()) && ss.getSheets().length > 1) {
      try { ss.deleteSheet(s); } catch(e) {}
    }
  });
}

// ── BAŞLIK YARDIMCILARI ──────────────────────────────────────
function _baslik(s, aralik, metin, renk, fontColor, fontSize) {
  const r = s.getRange(aralik);
  r.setValue(metin)
   .setFontSize(fontSize || 13)
   .setFontWeight("bold")
   .setBackground(renk || RENKLER.KOYU)
   .setFontColor(fontColor || RENKLER.BEYAZ)
   .setVerticalAlignment("middle")
   .setHorizontalAlignment("center");
}

function _sutunBasliklari(s, satir, sutunlar) {
  sutunlar.forEach((b, i) => {
    s.getRange(satir, i + 1)
     .setValue(b)
     .setFontWeight("bold")
     .setFontSize(9)
     .setBackground(RENKLER.BASLIK)
     .setFontColor(RENKLER.METIN)
     .setHorizontalAlignment("center")
     .setVerticalAlignment("middle")
     .setWrap(true);
  });
  s.setRowHeight(satir, 36);
}

function _validasyon(liste) {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(liste, true)
    .setAllowInvalid(false).build();
}

function _aralikValidasyon(ss, sayfa, sutun, aralikSutun) {
  const kaynak = ss.getSheetByName("⚙️ SETTINGS").getRange("A2:A100").getValues().filter(r => r[0]).map(r => r[0]);
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(kaynak, true)
    .setAllowInvalid(true).build();
}

// ── SİPARİŞLER SAYFASI ───────────────────────────────────────
function _kurSiparisler(ss) {
  const s = ss.getSheetByName("📋 ORDERS");
  s.setTabColor("#4285f4");

  // Başlık + buton alanı
  s.getRange("A1:L1").merge();
  _baslik(s, "A1:L1", "📋  ORDER TRACKING LIST", RENKLER.KOYU, RENKLER.BEYAZ, 14);
  s.setRowHeight(1, 45);

  // Buton satırı
  // Tıklanabilir buton hücresi
  s.getRange("A2:B2").merge()
   .setValue("  ➕  YENİ SİPARİŞ EKLE")
   .setBackground("#1a73e8").setFontColor("#ffffff")
   .setFontWeight("bold").setFontSize(11)
   .setHorizontalAlignment("center").setVerticalAlignment("middle")
   .setBorder(false,false,false,false,false,false);
  s.getRange("C2").setValue("← Tıklayın veya: Menu > Production Tracker > Yeni Sipariş Ekle")
   .setFontStyle("italic").setFontColor("#888888").setFontSize(9);
  s.setRowHeight(2, 30);

  // Sütun başlıkları
  _sutunBasliklari(s, 3, [
    "ORDER NO","ORDER DATE","MODEL","ORDER QTY",
    "PRODUCED","REMAINING","STATUS","ŞASE STATUSU","DELIVERY DATE","INVOICE","NOTES","LAST UPDATE"
  ]);
  s.setFrozenRows(3);

  // Kolon genişlikleri
  [100,110,240,120,100,100,120,140,110,100,200,130].forEach((w, i) => s.setColumnWidth(i + 1, w));

  // Dropdown validasyonlar
  // NOT: H ve diğer validasyonlar veri yazıldıktan sonra eklenir (_siparisleriAktar sonrası)
  s.getRange("G4:G500").setDataValidation(_validasyon(SIPARIS_STATUS));
  s.getRange("J4:J500").setDataValidation(_validasyon(["Not Issued","Issued","Partial"]));

  // Format
  s.getRange("B4:B500").setNumberFormat("dd.mm.yyyy");
  s.getRange("I4:I500").setNumberFormat("dd.mm.yyyy");
  s.getRange("L4:L500").setNumberFormat("dd.mm.yyyy hh:mm");
  s.getRange("D4:F500").setNumberFormat("#,##0");

  // Koşullu renk
  const ar = s.getRange("A4:L500");
  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$G4="Completed"').setBackground(RENKLER.YESIL).setRanges([ar]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$G4="In Production"').setBackground(RENKLER.MAVI).setRanges([ar]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$G4="Cancelled"').setBackground(RENKLER.KIRMIZI).setRanges([ar]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$G4="Waiting"').setBackground(RENKLER.SARI).setRanges([ar]).build(),
  ]);

  // Mevcut verileri aktar
  _siparisleriAktar(ss, s);
  // Veri yazıldıktan SONRA H sütunu validasyonu ekle
  s.getRange("H4:H500").setDataValidation(_validasyon(SASE_STATUS));
}

function _siparisleriAktar(ss, s) {
  // ── SAMPLE DATA — Replace with your own orders or leave empty ──
  // Orders are added via Menu → Add New Order after setup.
  // Uncomment and edit the lines below to pre-load sample data.

  const bugun = new Date();
  /*
  const ornek = [
    // [orderNo, "dd.mm.yyyy", "MODEL NAME", qty, produced, remaining, "Status", "Chassis Status"],
    [1, "01.01.2026", "MODEL-A 50CC",  500, 0, 500, "Waiting", "Not Ordered"],
    [1, "01.01.2026", "MODEL-B 125CC", 300, 0, 300, "Waiting", "Not Ordered"],
    [2, "15.01.2026", "MODEL-C ELECTRIC", 200, 200, 0, "Completed", "Fully Received"],
  ];
  const satirlar = ornek.map(r => {
    const d = (s) => new Date(s.split(".").reverse().join("-"));
    return [r[0], d(r[1]), r[2], r[3], r[4], r[5], r[6], r[7], "", r[4]===r[3]?"Issued":"Not Issued", "", bugun];
  });
  if (satirlar.length > 0) s.getRange(4, 1, satirlar.length, 12).setValues(satirlar);
  */
}

function _kurSaseTakip(ss) {
  const s = ss.getSheetByName("🔩 CHASSIS TRACKING");
  // Tüm eski içeriği, formatları ve validasyonları temizle
  s.clearContents();
  s.clearFormats();
  s.clearNotes();
  try { s.clearConditionalFormatRules(); } catch(e) {}
  // Tüm validasyonları sıfırla
  s.getRange("A1:T500").clearDataValidations();
  SpreadsheetApp.flush();
  s.setTabColor("#34a853");

  // Ana başlık
  s.getRange("A1:T1").merge();
  _baslik(s, "A1:T1", "🔩  CHASSIS TRACKING — BY SUPPLIER", RENKLER.KOYU, RENKLER.BEYAZ, 14);
  s.setRowHeight(1, 45);

  // Buton satırı
  s.getRange("A2:B2").merge()
   .setValue("  📦  ŞASE GİRİŞİ YAP")
   .setBackground("#34a853").setFontColor("#ffffff")
   .setFontWeight("bold").setFontSize(11)
   .setHorizontalAlignment("center").setVerticalAlignment("middle");
  s.getRange("C2:D2").merge()
   .setValue("  🔄  STATUS GÜNCELLE")
   .setBackground("#fbbc04").setFontColor("#333333")
   .setFontWeight("bold").setFontSize(11)
   .setHorizontalAlignment("center").setVerticalAlignment("middle");
  s.getRange("E2").setValue("← Tıklayın veya Menu > Production Tracker")
   .setFontStyle("italic").setFontColor("#888888").setFontSize(9);
  s.setRowHeight(2, 30);

  // ── ÖZET TABLO başlıkları (A:K) ──
  s.getRange("A3:K3").merge();
  _baslik(s, "A3:K3", "MODEL-BASED SUMMARY", RENKLER.BASLIK, RENKLER.METIN, 10);
  s.setRowHeight(3, 22);

  _sutunBasliklari(s, 4, [
    "MODEL","ORDER NO","ORDER QTY","SUPPLIER",
    "CHASSIS ORDER DATE","COMMITMENT DATE",
    "TOTAL RECEIVED","REMAINING","STATUS","LAST DELIVERY","NOT"
  ]);
  s.setFrozenRows(4);

  s.getRange("E5:F200").setNumberFormat("dd.mm.yyyy");
  s.getRange("J5:J200").setNumberFormat("dd.mm.yyyy");
  s.getRange("C5:G200").setNumberFormat("#,##0");
  s.getRange("H5:H200").setNumberFormat("#,##0");
  [240,100,120,140,110,110,110,110,140,110,200].forEach((w, i) => s.setColumnWidth(i + 1, w));

  // Koşullu renk - özet
  const ar = s.getRange("A5:K200");
  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$I5="Fully Received"').setBackground(RENKLER.YESIL).setRanges([ar]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$I5="Partial"').setBackground(RENKLER.SARI).setRanges([ar]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$I5="Ordered"').setBackground(RENKLER.MAVI).setRanges([ar]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$I5="Not Ordered"').setBackground(RENKLER.KIRMIZI).setRanges([ar]).build(),
  ]);
  // NOT: I5 validasyonu veri yazıldıktan SONRA ekleniyor (aşağıda)

  // Ayırıcı sütun
  s.getRange("L1:L500").setBackground("#cccccc");
  s.setColumnWidth(12, 8);

  // ── GİRİŞ LOGU başlıkları (M:T) ──
  s.getRange("M3:T3").merge();
  _baslik(s, "M3:T3", "CHASSIS ENTRY LOG (Each Delivery Separate Row)", RENKLER.BASLIK, RENKLER.METIN, 10);

  _sutunBasliklari(s, 4, [
    "","","","","","","","","","","",
    "","DELIVERY DATE","MODEL","SUPPLIER","WAYBILL NO","RECEIVED QTY","ORDER NO","RECORDED BY","NOT"
  ]);
  s.getRange("M5:M500").setNumberFormat("dd.mm.yyyy");
  s.getRange("Q5:Q500").setNumberFormat("#,##0");
  [110,200,140,130,110,100,100,200].forEach((w, i) => s.setColumnWidth(i + 13, w));

  // Mevcut verileri aktar
  function _donustur(v) {
    if (!v) return "Not Ordered";
    const e = v.toString().toUpperCase();
    if (e === "TAMAMI GELDİ" || e === "TAMAMI GELDI" || e === "TAMAMLANDI") return "Fully Received";
    if (e.includes("KISMI")) return "Partial";
    if (e.includes("VERİLDİ") || e.includes("VERILDI")) return "Ordered";
    return "Not Ordered";
  }

  const sipVerisi = ss.getSheetByName("📋 ORDERS").getRange("A4:H500").getValues();
  const satirlar = [];
  sipVerisi.forEach(r => {
    if (!r[0] || !r[2]) return;
    satirlar.push([r[2], r[0], r[3], "", "", "", r[4], r[5], _donustur(r[7]), "", ""]);
  });
  if (satirlar.length > 0) {
    const aralik = s.getRange(5, 1, satirlar.length, 11);
    aralik.clearDataValidations();
    aralik.setValues(satirlar);
    SpreadsheetApp.flush();
  }
  // Veri yazıldıktan SONRA validasyon ekle
  s.getRange("I5:I200").setDataValidation(_validasyon(SASE_STATUS));
}

// ── İŞ EMİRLERİ SAYFASI ──────────────────────────────────────
function _kurIsEmirleri(ss) {
  const s = ss.getSheetByName("⚙️ PRODUCTION ORDERS");
  s.setTabColor("#fbbc04");

  s.getRange("A1:N1").merge();
  _baslik(s, "A1:N1", "⚙️  ARP — PRODUCTION ORDERS", RENKLER.KOYU, RENKLER.BEYAZ, 14);
  s.setRowHeight(1, 45);
  s.getRange("A2:B2").merge()
   .setValue("  ⚙️  YENİ İŞ EMRİ AÇ")
   .setBackground("#f9ab00").setFontColor("#333333")
   .setFontWeight("bold").setFontSize(11)
   .setHorizontalAlignment("center").setVerticalAlignment("middle");
  s.getRange("C2:D2").merge()
   .setValue("  ✏️  İŞ EMRİ GÜNCELLE")
   .setBackground("#e37400").setFontColor("#ffffff")
   .setFontWeight("bold").setFontSize(11)
   .setHorizontalAlignment("center").setVerticalAlignment("middle");
  s.getRange("E2").setValue("← Tıklayın veya Menu > Production Tracker")
   .setFontStyle("italic").setFontColor("#888888").setFontSize(9);
  s.setRowHeight(2, 30);

  _sutunBasliklari(s, 3, [
    "ARP NO","ORDER NO","MODEL","BANT","PLANNED QTY",
    "START DATE","END DATE (PLAN)","STATUS","PRIORITY","DONE","REMAINING","STOP REASON","NOT","LAST UPDATE"
  ]);
  s.setFrozenRows(3);

  s.getRange("D4:D300").setDataValidation(_validasyon(LINES));
  s.getRange("H4:H300").setDataValidation(_validasyon(IE_STATUS));
  s.getRange("I4:I300").setDataValidation(_validasyon(ONCELIK));
  s.getRange("F4:G300").setNumberFormat("dd.mm.yyyy");
  s.getRange("N4:N300").setNumberFormat("dd.mm.yyyy hh:mm");
  s.getRange("E4:E300").setNumberFormat("#,##0");
  s.getRange("J4:K300").setNumberFormat("#,##0");

  // REMAINING sütunu formülsüz — isEmriGuncelle fonksiyonu günceller
  // Boş bırakıyoruz, hata olmasın

  const ar = s.getRange("A4:N300");
  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$H4="Completed"').setBackground(RENKLER.YESIL).setRanges([ar]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$H4="In Progress"').setBackground(RENKLER.MAVI).setRanges([ar]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$H4="Paused"').setBackground(RENKLER.SARI).setRanges([ar]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$H4="Planned"').setBackground(RENKLER.GRI).setRanges([ar]).build(),
  ]);

  [110,100,240,90,120,110,110,130,110,100,100,180,180,130].forEach((w, i) => s.setColumnWidth(i + 1, w));

  // ARP-0001 örnek veri
  s.getRange("A4:N4").setValues([["ARP-0001","34","POLO FARM","B BANT",1000,new Date("2026-07-04"),new Date("2026-07-04"),"Completed","3-Normal",1000,"","","",new Date()]]);
}

// ── BANT TAKİP SAYFASI ───────────────────────────────────────
function _kurBantTakip(ss) {
  const s = ss.getSheetByName("📈 LINE TRACKING");
  s.setTabColor("#ea4335");

  s.getRange("A1:H1").merge();
  _baslik(s, "A1:H1", "📈  DAILY LINE TRACKING", RENKLER.KOYU, RENKLER.BEYAZ, 14);
  s.setRowHeight(1, 45);
  s.getRange("A2:B2").merge()
   .setValue("  📈  BANT GİRİŞİ YAP")
   .setBackground("#ea4335").setFontColor("#ffffff")
   .setFontWeight("bold").setFontSize(11)
   .setHorizontalAlignment("center").setVerticalAlignment("middle");
  s.getRange("C2").setValue("← Tıklayın veya Menu > Production Tracker > Bant Girişi")
   .setFontStyle("italic").setFontColor("#888888").setFontSize(9);
  s.setRowHeight(2, 30);

  _sutunBasliklari(s, 3, ["TARİH","HAT","VEHICLE MODEL","RENK","PRODUCED ADET","TOTAL QTY","DESCRIPTION","ARP NO"]);
  s.setFrozenRows(3);

  s.getRange("B4:B500").setDataValidation(_validasyon(LINES));
  s.getRange("A4:A500").setNumberFormat("dd.mm.yyyy");
  s.getRange("E4:F500").setNumberFormat("#,##0");
  [110,100,240,80,120,120,200,110].forEach((w, i) => s.setColumnWidth(i + 1, w));

  const ar = s.getRange("A4:H500");
  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$G4="Bitti"').setBackground(RENKLER.YESIL).setRanges([ar]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$G4="End of day report"').setBackground(RENKLER.GRI).setRanges([ar]).build(),
  ]);
}

// ── DASHBOARD ─────────────────────────────────────────────────
function _kurDashboard(ss) {
  const s = ss.getSheetByName("📊 DASHBOARD");
  s.setTabColor("#4285f4");
  try { s.getRange("A1:M300").breakApart(); } catch(e) {}
  s.clearContents();
  s.clearFormats();
  SpreadsheetApp.flush();

  // Sütun genişlikleri - 12 sütun, G ayırıcı yok
  [280,80,100,100,100,120,120,120,120,120,120,140].forEach((w,i) => s.setColumnWidth(i+1,w));

  // ── BAŞLIK ──
  s.getRange("A1:L1").merge()
   .setValue("🏭  PRODUCTION TRACKER — OVERVIEW")
   .setFontSize(18).setFontWeight("bold")
   .setBackground("#4a6fa5").setFontColor("#ffffff")
   .setHorizontalAlignment("center").setVerticalAlignment("middle");
  s.setRowHeight(1, 55);

  // Son güncelleme
  s.getRange("A2").setValue("Last updated:").setFontColor("#999").setFontSize(9);
  s.getRange("B2:D2").merge().setValue(new Date()).setNumberFormat("dd.mm.yyyy hh:mm")
   .setFontColor("#999").setFontSize(9);
  s.getRange("I2:L2").merge()
   .setValue("🔄  Menu > Refresh Dashboard")
   .setFontStyle("italic").setFontColor("#aaaaaa").setFontSize(9)
   .setHorizontalAlignment("right");
  s.setRowHeight(2, 22);

  // ── KPI KARTLARI — 6 kart, 2 sütunluk ──
  const kpi = _dashKpiHesapla(ss);

  const kartBilgi = [
    ["A3:B3","A4:B4","📦  TOPLAM SİPARİŞ","#1a73e8","#ffffff", kpi.toplamSip],
    ["C3:D3","C4:D4","✅  COMPLETED",    "#34a853","#ffffff", kpi.tamamlanan],
    ["E3:F3","E4:F4","⚙️  IN PRODUCTION",      "#4285f4","#ffffff", kpi.uretimde],
    ["G3:H3","G4:H4","⏳  WAITING",      "#f9ab00","#333333", kpi.bekliyor],
    ["I3:J3","I4:J4","🔩  CHASSIS WAITING","#ea4335","#ffffff", kpi.saseBekleyen],
    ["K3:L3","K4:L4","📈  PROGRESS",      "#00897b","#ffffff", kpi.ilerleme+"%"],
  ];

  s.setRowHeight(3, 26);
  s.setRowHeight(4, 55);

  kartBilgi.forEach(([basAralik, degAralik, baslik, bg, fg, deger]) => {
    s.getRange(basAralik).merge()
     .setValue(baslik)
     .setBackground(bg).setFontColor(fg)
     .setFontSize(9).setFontWeight("bold")
     .setHorizontalAlignment("center").setVerticalAlignment("middle");

    s.getRange(degAralik).merge()
     .setValue(deger)
     .setBackground(bg).setFontColor(fg)
     .setFontSize(28).setFontWeight("bold")
     .setHorizontalAlignment("center").setVerticalAlignment("middle")
     .setNumberFormat(typeof deger === "number" ? "#,##0" : "@");
  });

  // Ayırıcı çizgi
  s.getRange("A5:L5").setBackground("#e8eaf6");
  s.setRowHeight(5, 6);

  // ── BANT ÖZETİ — Haftalık Tablo ──
  s.getRange("A6:L6").merge();
  _baslik(s, "A6:L6", "📊  WEEKLY LINE PRODUCTION (Last 7 Days)", "#16213e", "#ffffff", 11);
  s.setRowHeight(6, 28);

  // Gün başlıkları satırı
  const bugun = new Date(); bugun.setHours(0,0,0,0);
  const gunler = [];
  for (let g = 6; g >= 0; g--) {
    const d = new Date(bugun); d.setDate(d.getDate() - g); gunler.push(d);
  }

  const bantS2 = ss.getSheetByName("📈 LINE TRACKING");
  const bantVeri2 = bantS2.getRange("A4:E500").getValues();

  // Başlık satırı: boş | Gün1 | Gün2 | ... | Gün7 | TOPLAM
  const _tr = ["Paz","Pzt","Sal","Çar","Per","Cum","Cmt"];
  const gunBasliklari = ["LINE \ DAY"].concat(
    gunler.map(d => _tr[d.getDay()] + " " + Utilities.formatDate(d, Session.getScriptTimeZone(), "dd.MM"))
  ).concat(["WEEKLY TOTAL"]);

  // Sütun genişlikleri: A(bant adı) + B:H(7 gün) + I(toplam) = 9 sütun, L'ye kadar yay
  [110,120,120,120,120,120,120,120,140,60,60,60].forEach((w,i) => s.setColumnWidth(i+1, w));

  s.setRowHeight(7, 22);
  gunBasliklari.forEach((b, i) => {
    s.getRange(7, i+1).setValue(b)
     .setFontWeight("bold").setFontSize(9)
     .setBackground(i === 0 ? "#16213e" : (i === 8 ? "#1a1a2e" : "#2d3a8c"))
     .setFontColor("#ffffff")
     .setHorizontalAlignment("center").setVerticalAlignment("middle")
     .setWrap(true);
  });

  // Her bant için satır
  const bantColorler = ["#f5f5f5","#f5f5f5","#f5f5f5","#f5f5f5"];
  const bantKoyu   = ["#1a1a2e","#1a1a2e","#1a1a2e","#1a1a2e"];

  LINES.forEach((bant, bi) => {
    const satir = 8 + bi;
    s.setRowHeight(satir, 32);

    // Bant adı
    s.getRange(satir, 1).setValue(bant)
     .setFontWeight("bold").setFontSize(10)
     .setBackground("#16213e").setFontColor("#ffffff")
     .setHorizontalAlignment("center").setVerticalAlignment("middle");

    let haftaToplam = 0;
    gunler.forEach((gun, gi) => {
      let gunToplam = 0;
      bantVeri2.forEach(r => {
        if (!r[0] || r[1] !== bant) return;
        const t = new Date(r[0]); t.setHours(0,0,0,0);
        if (t.getTime() === gun.getTime()) gunToplam += Number(r[4]) || 0;
      });
      haftaToplam += gunToplam;

      const hucre = s.getRange(satir, gi+2);
      hucre.setValue(gunToplam > 0 ? gunToplam : "—")
       .setHorizontalAlignment("center").setVerticalAlignment("middle")
       .setFontSize(11).setFontWeight(gunToplam > 0 ? "bold" : "normal")
       .setBackground(gunToplam > 0 ? "#eef2ff" : "#fafafa")
       .setFontColor(gunToplam > 0 ? "#1a1a2e" : "#cccccc")
       .setNumberFormat(gunToplam > 0 ? "#,##0" : "@");
    });

    // Hafta toplamı
    s.getRange(satir, 9).setValue(haftaToplam)
     .setFontWeight("bold").setFontSize(13)
     .setBackground("#4a6fa5").setFontColor("#ffffff")
     .setHorizontalAlignment("center").setVerticalAlignment("middle")
     .setNumberFormat("#,##0");

    // Kalan sütunlar boş bırak
    s.getRange(satir, 10, 1, 3).setBackground("#ffffff");
  });

  // Genel toplam satırı
  s.setRowHeight(12, 28);
  s.getRange(12, 1).setValue("DAILY TOTAL")
   .setFontWeight("bold").setFontSize(9)
   .setBackground("#4a6fa5").setFontColor("#ffffff")
   .setHorizontalAlignment("center");

  gunler.forEach((gun, gi) => {
    let gunGenel = 0;
    bantVeri2.forEach(r => {
      if (!r[0]) return;
      const t = new Date(r[0]); t.setHours(0,0,0,0);
      if (t.getTime() === gun.getTime()) gunGenel += Number(r[4]) || 0;
    });
    s.getRange(12, gi+2).setValue(gunGenel > 0 ? gunGenel : "—")
     .setFontWeight("bold").setFontSize(10)
     .setBackground(gunGenel > 0 ? "#c5cae9" : "#f5f5f5")
     .setFontColor(gunGenel > 0 ? "#1a1a2e" : "#cccccc")
     .setHorizontalAlignment("center");
  });

  // Genel hafta toplamı
  let genelToplam = 0;
  bantVeri2.forEach(r => {
    if (!r[0]) return;
    const t = new Date(r[0]); t.setHours(0,0,0,0);
    if (t >= gunler[0]) genelToplam += Number(r[4]) || 0;
  });
  s.getRange(12, 9).setValue(genelToplam)
   .setFontWeight("bold").setFontSize(13)
   .setBackground("#4a6fa5").setFontColor("#ffffff")
   .setHorizontalAlignment("center").setNumberFormat("#,##0");
  s.getRange(12, 10, 1, 3).setBackground("#c5cae9");

  // Grafik varsa kaldır
  try { s.getCharts().forEach(c => s.removeChart(c)); } catch(e) {}

  // Ayırıcı
  s.getRange("A13:L13").setBackground("#e8eaf6");
  s.setRowHeight(13, 6);


  // ── SİPARİŞ TABLOSU BAŞLIK ──
  s.getRange("A14:L14").merge();
  _baslik(s, "A14:L14", "📋  ORDER DETAIL & PROGRESS", "#16213e", "#ffffff", 11);
  s.setRowHeight(14, 28);

  _sutunBasliklari(s, 15, [
    "MODEL","SİP NO","SİP ADET","PRODUCED","REMAINING","STATUS",
    "PROGRESS ÇUBUĞU","","","","ŞASE STATUSU","INVOICE"
  ]);
  s.getRange("G15:J15").merge().setValue("PROGRESS ÇUBUĞU")
   .setFontWeight("bold").setFontSize(9)
   .setBackground("#16213e").setFontColor("#e0e0e0")
   .setHorizontalAlignment("center").setVerticalAlignment("middle");
  s.setRowHeight(15, 30);
  s.setFrozenRows(15);

  s.getRange("A16").setValue("↻  Menu > Refresh Dashboard ile güncelleyin")
   .setFontStyle("italic").setFontColor("#aaaaaa").setFontSize(9);
}


function _kurDashboardBantTablosu(ss, s) {
  const bantS = ss.getSheetByName("📈 LINE TRACKING");
  const bantVeri = bantS.getRange("A4:E500").getValues();
  const bugun = new Date(); bugun.setHours(0,0,0,0);
  const gunler = [];
  for (let g = 6; g >= 0; g--) {
    const d = new Date(bugun); d.setDate(d.getDate() - g); gunler.push(d);
  }
  const bantColorler = ["#f5f5f5","#f5f5f5","#f5f5f5","#f5f5f5"];
  const bantKoyu   = ["#1a1a2e","#1a1a2e","#1a1a2e","#1a1a2e"];

  LINES.forEach((bant, bi) => {
    const satir = 8 + bi;
    let haftaToplam = 0;
    gunler.forEach((gun, gi) => {
      let gunToplam = 0;
      bantVeri.forEach(r => {
        if (!r[0] || r[1] !== bant) return;
        const t = new Date(r[0]); t.setHours(0,0,0,0);
        if (t.getTime() === gun.getTime()) gunToplam += Number(r[4]) || 0;
      });
      haftaToplam += gunToplam;
      s.getRange(satir, gi+2)
       .setValue(gunToplam > 0 ? gunToplam : "—")
       .setBackground(gunToplam > 0 ? bantColorler[bi] : "#f8f9fa")
       .setFontColor(gunToplam > 0 ? bantKoyu[bi] : "#cccccc")
       .setFontWeight(gunToplam > 0 ? "bold" : "normal");
    });
    s.getRange(satir, 9).setValue(haftaToplam)
     .setBackground(bantKoyu[bi]).setFontColor("#ffffff").setFontWeight("bold");
  });

  // Genel toplam satırı
  gunler.forEach((gun, gi) => {
    let gunGenel = 0;
    bantVeri.forEach(r => {
      if (!r[0]) return;
      const t = new Date(r[0]); t.setHours(0,0,0,0);
      if (t.getTime() === gun.getTime()) gunGenel += Number(r[4]) || 0;
    });
    s.getRange(12, gi+2)
     .setValue(gunGenel > 0 ? gunGenel : "—")
     .setBackground(gunGenel > 0 ? "#c5cae9" : "#f5f5f5")
     .setFontColor(gunGenel > 0 ? "#1a1a2e" : "#cccccc");
  });
  let genelToplam = 0;
  bantVeri.forEach(r => {
    if (!r[0]) return;
    const t = new Date(r[0]); t.setHours(0,0,0,0);
    if (t >= gunler[0]) genelToplam += Number(r[4]) || 0;
  });
  s.getRange(12, 9).setValue(genelToplam).setBackground("#3d5a99").setFontColor("#ffffff").setFontWeight("bold");
}

function _dashKpiHesapla(ss) {
  const sipVeri = ss.getSheetByName("📋 ORDERS").getRange("A4:J500").getValues();
  const saseVeri = ss.getSheetByName("🔩 CHASSIS TRACKING").getRange("I5:I200").getValues();
  let toplamSip=0, tamamlanan=0, uretimde=0, bekliyor=0, toplamUretilen=0;
  sipVeri.forEach(r => {
    if (!r[0] || !r[2]) return;
    const durum = r[6].toString();
    const adet = Number(r[3]) || 0;
    if (durum === "Cancelled") return;
    toplamSip += adet;
    toplamUretilen += Number(r[4]) || 0;
    if (durum === "Completed") tamamlanan += adet;
    else if (durum === "In Production") uretimde += adet;
    else if (durum === "Waiting") bekliyor += adet;
  });
  let saseBekleyen = 0;
  saseVeri.forEach(r => { if (r[0] === "Not Ordered") saseBekleyen++; });
  const ilerleme = toplamSip > 0 ? Math.round((toplamUretilen / toplamSip) * 100) : 0;
  return { toplamSip, tamamlanan, uretimde, bekliyor, saseBekleyen, ilerleme };
}

// ── AYARLAR SAYFASI ───────────────────────────────────────────
function _kurAyarlar(ss) {
  const s = ss.getSheetByName("⚙️ SETTINGS");

  _baslik(s, "A1", "MODELS", RENKLER.KOYU, RENKLER.BEYAZ, 11);
  _baslik(s, "C1", "SUPPLIERLER", RENKLER.KOYU, RENKLER.BEYAZ, 11);
  _baslik(s, "E1", "LINES", RENKLER.KOYU, RENKLER.BEYAZ, 11);

  const modeller = [
    "AR 50 (CG 50)","AR50000 YENİ","AR-250","BLADE ELEKTRİKLİ ATV",
    "BRAVO 50","BRAVO 50CC","BRAVO 125","BRAVO 125CC",
    "DERYA PRO","DERYA NEW YENİ 3 TEKER","DAZZLE PRO",
    "E-KARGO","E-PICK UP YENİ","E-PICK UP BIG YENİ","E-PICK UP SMALL YENİ",
    "FELIX NEW","FELIX PRO","FREEDOM 50CC","FIRTINA 50CC",
    "GALAXY YENİ","GOLF 01 YENİ",
    "KABİNLİ ÜÇ TEKER DOZ","KABİNSİZ ÜÇ TEKER DOZ",
    "MAX PRO 150","MAX PRO 250","MAX T NEW YENİ",
    "MINI CARGO","MINI CARGO YENİ","MINI FELIX YENİ","MT 125",
    "NAVARA NEW","NAVARA NEW X YENİ","NOSTALJİ YENİ",
    "POLO FARM","POLO FARM NEW YENİ","POLO PLUS","POLO PLUS PRO",
    "PRIME 125 (50)","RAPTOR ATV","RUZGAR NEW","RUZGAR PRO",
    "S1","S2 YENİ","SONİC 125 BENZİNLİ ATV","SUPER 50",
    "ÜÇ TEKER BENZİNLİ (KABİNLİ)","ÜÇ TEKER BENZİNLİ (KABİNLİ) YENİ",
    "ÜÇ TEKER DOZ","ÜÇ TEKER DOZ KABİNLİ","ÜÇ TEKER DOZ KABİNLİ (AR10000 NEW)",
    "ZR-6","ZR3","ZR4 YENİ"
  ];
  const tedarikciler = ["Antep Supplier","Tarsus Supplier"];

  s.getRange(2, 1, modeller.length, 1).setValues(modeller.map(m => [m]));
  s.getRange(2, 3, tedarikciler.length, 1).setValues(tedarikciler.map(t => [t]));
  LINES.forEach((b, i) => s.getRange(i + 2, 5).setValue(b));
  s.setColumnWidth(1, 280);
  s.setColumnWidth(3, 180);
  s.setColumnWidth(5, 120);

  // Ek bilgi
  s.getRange("A60").setValue("⬆️ Yeni model eklemek için yukarıdaki listeye yazın.")
   .setFontStyle("italic").setFontColor("#888888").setFontSize(9);
  s.getRange("C20").setValue("⬆️ Yeni tedarikçi ekleyin.")
   .setFontStyle("italic").setFontColor("#888888").setFontSize(9);
}

// ── MENÜ ─────────────────────────────────────────────────────
function onEdit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row !== 2) return;

  const ad = sheet.getName();
  if (ad === "📋 ORDERS" && col <= 2) {
    e.range.setValue("  ➕  YENİ SİPARİŞ EKLE");
    siparisEkleSidebar();
  } else if (ad === "🔩 CHASSIS TRACKING" && col <= 2) {
    e.range.setValue("  📦  ŞASE GİRİŞİ YAP");
    saseGirisiSidebar();
  } else if (ad === "🔩 CHASSIS TRACKING" && col >= 3 && col <= 4) {
    e.range.getSheet().getRange("C2:D2").setValue("  🔄  STATUS GÜNCELLE");
    saseDurumSidebar();
  } else if (ad === "⚙️ PRODUCTION ORDERS" && col <= 2) {
    e.range.setValue("  ⚙️  YENİ İŞ EMRİ AÇ");
    isEmriSidebar();
  } else if (ad === "⚙️ PRODUCTION ORDERS" && col >= 3 && col <= 4) {
    e.range.getSheet().getRange("C2:D2").setValue("  ✏️  İŞ EMRİ GÜNCELLE");
    isEmriGuncelleSidebar();
  } else if (ad === "📈 LINE TRACKING" && col <= 2) {
    e.range.setValue("  📈  BANT GİRİŞİ YAP");
    bantGirisiSidebar();
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🏭 Production Tracker")
    .addItem("📋 Add New Order","siparisEkleSidebar")
    .addSeparator()
    .addItem("🔩 Chassis Entry","saseGirisiSidebar")
    .addItem("🔩 Update Chassis Status","saseDurumSidebar")
    .addSeparator()
    .addItem("⚙️ Create New ARP","isEmriSidebar")
    .addItem("⚙️ Update ARP","isEmriGuncelleSidebar")
    .addSeparator()
    .addItem("📈 Line Entry","bantGirisiSidebar")
    .addSeparator()
    .addItem("🔄 Refresh Dashboard","dashboardYenile")
    .addSeparator()
    .addItem("⚙️ Settings (Add Model/Supplier)","ayarlariGoster")
    .addSeparator()
    .addItem("🧹 Fix Production Orders Errors","isEmirleriDuzelt")
    .addSeparator()
    .addItem("📥 Backup (Save to Drive)","yedekAl")
    .addItem("🧹 Archive & Clear Line Tracking","bantArsivle")
    .addSeparator()
    .addItem("🔧 Setup / Reset System (Admin)","kurulumKorumal")
    .addToUi();
}

// ── SİDEBAR HTML YARDIMCISI ──────────────────────────────────
function _sidebarAc(baslik, icerik) {
  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 13px;
         background: #f8f9fa; color: #333; padding: 12px; }
  h2 { background: #1a1a2e; color: white; padding: 12px 14px;
       margin: -12px -12px 16px; font-size: 14px; letter-spacing: 0.3px; }
  .form-group { margin-bottom: 12px; }
  label { display: block; font-weight: 600; margin-bottom: 4px;
          color: #444; font-size: 12px; }
  select, input[type=text], input[type=number], input[type=date], textarea {
    width: 100%; padding: 7px 9px; border: 1px solid #ccc;
    border-radius: 5px; font-size: 13px; background: white;
    transition: border-color 0.2s; }
  select:focus, input:focus, textarea:focus {
    outline: none; border-color: #4285f4; }
  textarea { resize: vertical; min-height: 60px; }
  .btn { display: block; width: 100%; padding: 10px;
         border: none; border-radius: 6px; font-size: 13px;
         font-weight: 600; cursor: pointer; margin-top: 6px; }
  .btn-primary { background: #1a1a2e; color: white; }
  .btn-primary:hover { background: #16213e; }
  .btn-secondary { background: #e9ecef; color: #333; }
  .msg { padding: 10px 12px; border-radius: 5px; margin-top: 10px;
         font-size: 12px; display: none; }
  .msg.ok { background: #d4edda; color: #155724; display: block; }
  .msg.err { background: #f8d7da; color: #721c24; display: block; }
  .divider { border: none; border-top: 1px solid #dee2e6; margin: 14px 0; }
  .hint { font-size: 11px; color: #888; margin-top: 3px; }
</style>
</head>
<body>
<h2>${baslik}</h2>
${icerik}
</body>
</html>`)
    .setTitle(baslik)
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── SİPARİŞ EKLEme SİDEBARI ─────────────────────────────────
function siparisEkleSidebar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const modeller = ss.getSheetByName("⚙️ SETTINGS").getRange("A2:A100").getValues()
    .filter(r => r[0]).map(r => `<option>${r[0]}</option>`).join("");

  // Mevcut sipariş numaralarını bul (tekrarsız, azalan sırada)
  const sipSayfa = ss.getSheetByName("📋 ORDERS");
  const noslar = sipSayfa.getRange("A4:A500").getValues()
    .filter(r => r[0]).map(r => Number(r[0]));
  const sonNo = noslar.length > 0 ? Math.max(...noslar) + 1 : 38;
  const mevcutNolar = [...new Set(noslar)].sort((a,b) => b-a);
  const mevcutOptions = mevcutNolar.map(n => `<option value="${n}">SİP-${n}</option>`).join("");

  _sidebarAc("📋 Add New Order", `
<div class="form-group">
  <label>Order Type *</label>
  <select id="sipTip" onchange="tipDegisti(this.value)">
    <option value="yeni">🆕 New Order (SİP-${sonNo})</option>
    <option value="mevcut">➕ Add Item to Existing Order</option>
  </select>
</div>
<div id="mevcutDiv" style="display:none" class="form-group">
  <label>Which Order? *</label>
  <select id="mevcutNo">
    <option value="">— Seçin —</option>${mevcutOptions}
  </select>
</div>
<div class="form-group">
  <label>Model *</label>
  <select id="model"><option value="">— Seçin —</option>${modeller}</select>
</div>
<div class="form-group">
  <label>Order Quantity *</label>
  <input type="number" id="adet" min="1" placeholder="Örn: 1000">
</div>
<div id="tarihDiv" class="form-group">
  <label>Order Date</label>
  <input type="date" id="tarih" value="${new Date().toISOString().split('T')[0]}">
</div>
<div class="form-group">
  <label>Delivery Date (optional)</label>
  <input type="date" id="teslim">
</div>
<div class="form-group">
  <label>Not</label>
  <textarea id="not" placeholder="Add if applicable..."></textarea>
</div>
<button class="btn btn-primary" onclick="kaydet()">✅ Kaydet</button>
<button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
<div id="msg" class="msg"></div>
<script>
function tipDegisti(val) {
  document.getElementById('mevcutDiv').style.display = val === 'mevcut' ? 'block' : 'none';
  document.getElementById('tarihDiv').style.display = val === 'yeni' ? 'block' : 'none';
}
function kaydet() {
  const tip = document.getElementById('sipTip').value;
  const model = document.getElementById('model').value;
  const adet = document.getElementById('adet').value;
  if (!model) { goster('Model seçin!', false); return; }
  if (!adet || adet < 1) { goster('Geçerli adet girin!', false); return; }
  let sipNo, tarih;
  if (tip === 'mevcut') {
    sipNo = document.getElementById('mevcutNo').value;
    if (!sipNo) { goster('Sipariş numarası seçin!', false); return; }
    tarih = '';
  } else {
    sipNo = '${sonNo}';
    tarih = document.getElementById('tarih').value;
  }
  document.querySelector('.btn-primary').disabled = true;
  document.querySelector('.btn-primary').textContent = 'Saving...';
  google.script.run
    .withSuccessHandler(r => goster(r, true))
    .withFailureHandler(e => goster(e.message, false))
    .siparisKaydet(
      sipNo, model, parseInt(adet), tarih,
      document.getElementById('teslim').value,
      document.getElementById('not').value
    );
}
function goster(m, ok) {
  const d = document.getElementById('msg');
  d.className = 'msg ' + (ok ? 'ok' : 'err');
  d.textContent = m;
  if (ok) setTimeout(() => google.script.host.close(), 2000);
  else {
    document.querySelector('.btn-primary').disabled = false;
    document.querySelector('.btn-primary').textContent = '✅ Kaydet';
  }
}
</script>`);
}

function siparisKaydet(sipNo, model, adet, tarih, teslim, not) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName("📋 ORDERS");

  // Son dolu satırı bul
  const veriler = s.getRange("A4:A500").getValues();
  let sonSatir = 4;
  for (let i = veriler.length - 1; i >= 0; i--) {
    if (veriler[i][0] !== "") { sonSatir = i + 5; break; }
  }

  // Mevcut siparişe ekleniyorsa tarihi o siparişten al
  let kayitDate;
  if (tarih) {
    kayitDate = new Date(tarih);
  } else {
    // Aynı sipariş nosunu bul, tarihini al
    const tumVeri = s.getRange("A4:B500").getValues();
    const eslesen = tumVeri.find(r => Number(r[0]) === Number(sipNo));
    kayitDate = eslesen && eslesen[1] ? eslesen[1] : new Date();
  }

  const tl = teslim ? new Date(teslim) : "";
  s.getRange(sonSatir, 1).setValue(Number(sipNo));
  s.getRange(sonSatir, 2).setValue(kayitDate);
  s.getRange(sonSatir, 3).setValue(model);
  s.getRange(sonSatir, 4).setValue(adet);
  s.getRange(sonSatir, 5).setValue(0);
  s.getRange(sonSatir, 6).setValue(adet);
  s.getRange(sonSatir, 7).setValue("Waiting");
  s.getRange(sonSatir, 8).setValue("Not Ordered");
  if (tl) s.getRange(sonSatir, 9).setValue(tl);
  s.getRange(sonSatir, 10).setValue("Not Issued");
  s.getRange(sonSatir, 11).setValue(not || "");
  s.getRange(sonSatir, 12).setValue(new Date());

  // Şase sayfasına da ekle
  const sase = ss.getSheetByName("🔩 CHASSIS TRACKING");
  const saseVeri = sase.getRange("A4:A200").getValues();
  let saseSon = 4;
  for (let i = saseVeri.length - 1; i >= 0; i--) {
    if (saseVeri[i][0] !== "") { saseSon = i + 5; break; }
  }
  sase.getRange(saseSon, 1).setValue(model);
  sase.getRange(saseSon, 2).setValue(Number(sipNo));
  sase.getRange(saseSon, 3).setValue(adet);
  sase.getRange(saseSon, 7).setValue(0);
  sase.getRange(saseSon, 8).setValue(adet);
  sase.getRange(saseSon, 9).setValue("Not Ordered");

  logYaz("SİPARİŞ", "YENİ KALEM", `SIP-${sipNo} | ${model} | ${adet} adet`);
  return `✅ Saved! SIP-${sipNo} — ${model} (${adet} adet)`;
}

// ── ŞASE GİRİŞİ SİDEBARI ─────────────────────────────────────
function saseGirisiSidebar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const modeller = ss.getSheetByName("⚙️ SETTINGS").getRange("A2:A100").getValues()
    .filter(r => r[0]).map(r => `<option>${r[0]}</option>`).join("");
  const tedarikciler = ss.getSheetByName("⚙️ SETTINGS").getRange("C2:C50").getValues()
    .filter(r => r[0]).map(r => `<option>${r[0]}</option>`).join("");
  const sipNolar = [...new Set(
    ss.getSheetByName("📋 ORDERS").getRange("A4:A500").getValues()
      .filter(r => r[0]).map(r => r[0])
  )].sort((a,b) => b-a).map(n => `<option value="${n}">SİP-${n}</option>`).join("");

  _sidebarAc("🔩 Şase Girişi", `
<div class="form-group">
  <label>Model *</label>
  <select id="model"><option value="">— Seçin —</option>${modeller}</select>
</div>
<div class="form-group">
  <label>Sipariş No *</label>
  <select id="sipNo"><option value="">— Seçin —</option>${sipNolar}</select>
</div>
<div class="form-group">
  <label>Supplier *</label>
  <select id="tedarikci"><option value="">— Seçin —</option>${tedarikciler}</select>
</div>
<div class="form-group">
  <label>Şase Order Date</label>
  <input type="date" id="sipDate">
  <div class="hint">Date order was placed with supplier</div>
</div>
<div class="form-group">
  <label>Taahhüt Datei</label>
  <input type="date" id="taahhutDate">
  <div class="hint">Date supplier committed to deliver</div>
</div>
<hr style="border:none;border-top:1px solid #dee2e6;margin:12px 0">
<div class="form-group">
  <label>Arrival Date *</label>
  <input type="date" id="gelisDate" value="${new Date().toISOString().split('T')[0]}">
  <div class="hint">Date chassis physically arrived</div>
</div>
<div class="form-group">
  <label>İrsaliye No</label>
  <input type="text" id="irsaliye" placeholder="Örn: IRSALİYE-2026-001">
</div>
<div class="form-group">
  <label>Gelen Adet *</label>
  <input type="number" id="adet" min="1" placeholder="How many units arrived in this shipment?">
</div>
<div class="form-group">
  <label>Not</label>
  <textarea id="not" placeholder="Add if applicable..."></textarea>
</div>
<button class="btn btn-primary" onclick="kaydet()">✅ Kaydet</button>
<button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
<div id="msg" class="msg"></div>
<script>
function kaydet() {
  const model = document.getElementById('model').value;
  const sipNo = document.getElementById('sipNo').value;
  const ted = document.getElementById('tedarikci').value;
  const adet = document.getElementById('adet').value;
  const gelis = document.getElementById('gelisDate').value;
  if (!model) { goster('Model seçin!', false); return; }
  if (!sipNo) { goster('Sipariş no seçin!', false); return; }
  if (!ted) { goster('Tedarikçi seçin!', false); return; }
  if (!adet || adet < 1) { goster('Geçerli adet girin!', false); return; }
  if (!gelis) { goster('Geliş tarihi girin!', false); return; }
  document.querySelector('.btn-primary').disabled = true;
  document.querySelector('.btn-primary').textContent = 'Saving...';
  google.script.run
    .withSuccessHandler(r => goster(r, true))
    .withFailureHandler(e => goster(e.message, false))
    .saseKaydet(
      model, sipNo, ted,
      document.getElementById('sipDate').value,
      document.getElementById('taahhutDate').value,
      gelis,
      document.getElementById('irsaliye').value,
      parseInt(adet),
      document.getElementById('not').value
    );
}
function goster(m, ok) {
  const d = document.getElementById('msg');
  d.className = 'msg ' + (ok ? 'ok' : 'err');
  d.textContent = m;
  if (ok) setTimeout(() => google.script.host.close(), 2500);
  else {
    document.querySelector('.btn-primary').disabled = false;
    document.querySelector('.btn-primary').textContent = '✅ Kaydet';
  }
}
</script>`);
}

function saseKaydet(model, sipNo, tedarikci, sipDate, taahhutDate, gelisDate, irsaliye, adet, not) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName("🔩 CHASSIS TRACKING");

  // ── 1. ÖZET TABLOYU GÜNCELLE (A:K) ──
  const ozetVeri = s.getRange("A5:K200").getValues();
  let ozetSatir = -1;
  for (let i = 0; i < ozetVeri.length; i++) {
    if (ozetVeri[i][0].toString().toUpperCase() === model.toUpperCase() &&
        Number(ozetVeri[i][1]) === Number(sipNo)) {
      ozetSatir = i + 5;
      break;
    }
  }

  const gelisDateObj = gelisDate ? new Date(gelisDate) : new Date();

  if (ozetSatir > 0) {
    // Mevcut satırı güncelle
    const mevcutGelen = Number(s.getRange(ozetSatir, 7).getValue()) || 0;
    const sipAdet = Number(s.getRange(ozetSatir, 3).getValue()) || 0;
    const yeniGelen = mevcutGelen + adet;
    const kalan = Math.max(0, sipAdet - yeniGelen);
    const durum = yeniGelen >= sipAdet && sipAdet > 0 ? "Fully Received" : "Partial";

    s.getRange(ozetSatir, 4).setValue(tedarikci);
    if (sipDate) s.getRange(ozetSatir, 5).setValue(new Date(sipDate));
    if (taahhutDate) s.getRange(ozetSatir, 6).setValue(new Date(taahhutDate));
    s.getRange(ozetSatir, 7).setValue(yeniGelen);
    s.getRange(ozetSatir, 8).setValue(kalan);
    s.getRange(ozetSatir, 9).setValue(durum);
    s.getRange(ozetSatir, 10).setValue(gelisDateObj);
    if (not) s.getRange(ozetSatir, 11).setValue(not);
  } else {
    // Yeni satır ekle (sipariş listede yok)
    const sonBos = ozetVeri.findIndex(r => !r[0]) + 5;
    s.getRange(sonBos, 1).setValue(model);
    s.getRange(sonBos, 2).setValue(Number(sipNo));
    s.getRange(sonBos, 3).setValue(adet);
    s.getRange(sonBos, 4).setValue(tedarikci);
    if (sipDate) s.getRange(sonBos, 5).setValue(new Date(sipDate));
    if (taahhutDate) s.getRange(sonBos, 6).setValue(new Date(taahhutDate));
    s.getRange(sonBos, 7).setValue(adet);
    s.getRange(sonBos, 8).setValue(0);
    s.getRange(sonBos, 9).setValue("Partial");
    s.getRange(sonBos, 10).setValue(gelisDateObj);
    if (not) s.getRange(sonBos, 11).setValue(not);
  }

  // ── 2. GİRİŞ LOGUNA EKLE (M:T) — her teslimat ayrı satır ──
  const logVeri = s.getRange("M5:M500").getValues();
  let logSon = 5;
  for (let i = logVeri.length - 1; i >= 0; i--) {
    if (logVeri[i][0] !== "") { logSon = i + 6; break; }
  }
  s.getRange(logSon, 13).setValue(gelisDateObj);  // M - geliş tarihi
  s.getRange(logSon, 14).setValue(model);            // N - model
  s.getRange(logSon, 15).setValue(tedarikci);        // O - tedarikçi
  s.getRange(logSon, 16).setValue(irsaliye || "");   // P - irsaliye
  s.getRange(logSon, 17).setValue(adet);             // Q - gelen adet
  s.getRange(logSon, 18).setValue(Number(sipNo));    // R - sipariş no
  s.getRange(logSon, 19).setValue(Session.getActiveUser().getEmail() || "-"); // S
  s.getRange(logSon, 20).setValue(not || "");        // T - not

  // ── 3. SİPARİŞ SAYFASINI GÜNCELLE ──
  const sipS = ss.getSheetByName("📋 ORDERS");
  const sipVeri = sipS.getRange("A4:H500").getValues();
  for (let i = 0; i < sipVeri.length; i++) {
    if (sipVeri[i][2].toString().toUpperCase() === model.toUpperCase() &&
        Number(sipVeri[i][0]) === Number(sipNo)) {
      const sipAdet = Number(sipVeri[i][3]) || 0;
      const mevcutGelen = Number(s.getRange(ozetSatir > 0 ? ozetSatir : 5, 7).getValue()) || 0;
      const dur = mevcutGelen >= sipAdet && sipAdet > 0 ? "Fully Received" : "Partial";
      sipS.getRange(i + 4, 8).setValue(dur);
      sipS.getRange(i + 4, 12).setValue(new Date());
      break;
    }
  }

  logYaz("ŞASE", "GİRİŞ", `${model} | SIP-${sipNo} | ${tedarikci} | ${irsaliye} | ${adet} adet`);
  return `✅ Saved! ${model} — ${adet} adet (SIP-${sipNo})\nGiriş loguna eklendi.`;
}

// ── ŞASE STATUS GÜNCELLE SİDEBARI ─────────────────────────────
function saseDurumSidebar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const modeller = ss.getSheetByName("🔩 CHASSIS TRACKING").getRange("A4:A200").getValues()
    .filter(r => r[0]).map(r => `<option>${r[0]}</option>`).join("");
  const durumlar = SASE_STATUS.map(d => `<option>${d}</option>`).join("");

  _sidebarAc("🔩 Şase Durum Güncelle", `
<div class="form-group">
  <label>Model *</label>
  <select id="model"><option value="">— Seçin —</option>${modeller}</select>
</div>
<div class="form-group">
  <label>Yeni Durum *</label>
  <select id="durum"><option value="">— Seçin —</option>${durumlar}</select>
</div>
<div class="form-group">
  <label>Not</label>
  <textarea id="not" placeholder="Add description..."></textarea>
</div>
<button class="btn btn-primary" onclick="kaydet()">✅ Güncelle</button>
<button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
<div id="msg" class="msg"></div>
<script>
function kaydet() {
  const model = document.getElementById('model').value;
  const durum = document.getElementById('durum').value;
  if (!model || !durum) { goster('Model ve durum seçin!', false); return; }
  document.querySelector('.btn-primary').disabled = true;
  google.script.run
    .withSuccessHandler(r => goster(r, true))
    .withFailureHandler(e => goster(e.message, false))
    .saseDurumGuncelle(model, durum, document.getElementById('not').value);
}
function goster(m, ok) {
  const d = document.getElementById('msg');
  d.className = 'msg ' + (ok ? 'ok' : 'err');
  d.textContent = m;
  if (ok) setTimeout(() => google.script.host.close(), 2000);
  else { document.querySelector('.btn-primary').disabled = false; }
}
</script>`);
}

function saseDurumGuncelle(model, durum, not) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName("🔩 CHASSIS TRACKING");
  const veriler = s.getRange("A4:A200").getValues();
  for (let i = 0; i < veriler.length; i++) {
    if (veriler[i][0].toString().toUpperCase() === model.toUpperCase()) {
      s.getRange(i + 4, 9).setValue(durum);
      if (not) s.getRange(i + 4, 11).setValue(not);
      s.getRange(i + 4, 10).setValue(new Date());
      logYaz("ŞASE", "GÜNCELLEME", `${model} → ${durum}`);
      return `✅ ${model} durumu güncellendi: ${durum}`;
    }
  }
  return "⚠️ Model bulunamadı.";
}

// ── İŞ EMRİ SİDEBARI ─────────────────────────────────────────
function isEmriSidebar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const modeller = ss.getSheetByName("⚙️ SETTINGS").getRange("A2:A100").getValues()
    .filter(r => r[0]).map(r => `<option>${r[0]}</option>`).join("");
  const bantlar = LINES.map(b => `<option>${b}</option>`).join("");
  const oncelikler = ONCELIK.map(o => `<option>${o}</option>`).join("");

  // Sipariş noları
  const sipNolar = [...new Set(
    ss.getSheetByName("📋 ORDERS").getRange("A4:A500").getValues()
      .filter(r => r[0]).map(r => r[0])
  )].sort((a,b) => b-a).map(n => `<option>SIP-${n}</option>`).join("");

  // Son IE no
  const ieVeriler = ss.getSheetByName("⚙️ PRODUCTION ORDERS").getRange("A4:A300").getValues()
    .filter(r => r[0] && r[0].toString().startsWith("ARP-"));
  const sonIE = ieVeriler.length > 0
    ? Math.max(...ieVeriler.map(r => parseInt(r[0].toString().replace("ARP-","")) || 0)) + 1
    : 2;
  const ieNo = "ARP-" + String(sonIE).padStart(4, "0");

  _sidebarAc("⚙️ Yeni ARP", `
<div class="form-group">
  <label>ARP No</label>
  <input type="text" id="ieNo" value="${ieNo}" readonly style="background:#eef;font-weight:bold;">
</div>
<div class="form-group">
  <label>Sipariş No</label>
  <select id="sipNo"><option value="">— Seçin —</option>${sipNolar}</select>
</div>
<div class="form-group">
  <label>Model *</label>
  <select id="model"><option value="">— Seçin —</option>${modeller}</select>
</div>
<div class="form-group">
  <label>Line *</label>
  <select id="bant"><option value="">— Seçin —</option>${bantlar}</select>
</div>
<div class="form-group">
  <label>Planned Quantity *</label>
  <input type="number" id="adet" min="1" placeholder="How many units to produce?">
</div>
<div class="form-group">
  <label>Start Date</label>
  <input type="date" id="baslama" value="${new Date().toISOString().split('T')[0]}">
</div>
<div class="form-group">
  <label>End Date (Planned)</label>
  <input type="date" id="bitis">
</div>
<div class="form-group">
  <label>Priority</label>
  <select id="oncelik">${oncelikler}</select>
</div>
<div class="form-group">
  <label>Not</label>
  <textarea id="not" placeholder="Varsa not ekleyin..."></textarea>
</div>
<button class="btn btn-primary" onclick="kaydet()">✅ ARP Aç</button>
<button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
<div id="msg" class="msg"></div>
<script>
function kaydet() {
  const model = document.getElementById('model').value;
  const bant = document.getElementById('bant').value;
  const adet = document.getElementById('adet').value;
  if (!model) { goster('Model seçin!', false); return; }
  if (!bant) { goster('Bant seçin!', false); return; }
  if (!adet || adet < 1) { goster('Geçerli adet girin!', false); return; }
  document.querySelector('.btn-primary').disabled = true;
  document.querySelector('.btn-primary').textContent = 'Saving...';
  google.script.run
    .withSuccessHandler(r => goster(r, true))
    .withFailureHandler(e => goster(e.message, false))
    .isEmriKaydet(
      document.getElementById('ieNo').value,
      document.getElementById('sipNo').value,
      model, bant, parseInt(adet),
      document.getElementById('baslama').value,
      document.getElementById('bitis').value,
      document.getElementById('oncelik').value,
      document.getElementById('not').value
    );
}
function goster(m, ok) {
  const d = document.getElementById('msg');
  d.className = 'msg ' + (ok ? 'ok' : 'err');
  d.textContent = m;
  if (ok) setTimeout(() => google.script.host.close(), 2000);
  else { document.querySelector('.btn-primary').disabled = false;
         document.querySelector('.btn-primary').textContent = '✅ ARP Aç'; }
}
</script>`);
}

function isEmriKaydet(ieNo, sipNo, model, bant, adet, baslama, bitis, oncelik, not) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName("⚙️ PRODUCTION ORDERS");
  const veriler = s.getRange("A4:A300").getValues();
  let sonSatir = 4;
  for (let i = veriler.length - 1; i >= 0; i--) {
    if (veriler[i][0] !== "") { sonSatir = i + 5; break; }
  }
  s.getRange(sonSatir, 1).setValue(ieNo);
  s.getRange(sonSatir, 2).setValue(sipNo || "");
  s.getRange(sonSatir, 3).setValue(model);
  s.getRange(sonSatir, 4).setValue(bant);
  s.getRange(sonSatir, 5).setValue(adet);
  if (baslama) s.getRange(sonSatir, 6).setValue(new Date(baslama));
  if (bitis) s.getRange(sonSatir, 7).setValue(new Date(bitis));
  s.getRange(sonSatir, 8).setValue("Planned");
  s.getRange(sonSatir, 9).setValue(oncelik || "3-Normal");
  s.getRange(sonSatir, 10).setValue(0);
  s.getRange(sonSatir, 13).setValue(not || "");
  s.getRange(sonSatir, 14).setValue(new Date());

  // Sipariş sayfasında durumu Üretimde yap
  if (sipNo) {
    const sipS = ss.getSheetByName("📋 ORDERS");
    const sipVeri = sipS.getRange("C4:G500").getValues();
    for (let i = 0; i < sipVeri.length; i++) {
      if (sipVeri[i][0].toString().toUpperCase() === model.toUpperCase()) {
        if (sipVeri[i][4] === "Waiting") {
          sipS.getRange(i + 4, 7).setValue("In Production");
          sipS.getRange(i + 4, 12).setValue(new Date());
        }
        break;
      }
    }
  }
  logYaz("İŞ EMRİ", "YENİ", `${ieNo} | ${model} | ${bant} | ${adet} adet`);
  return `✅ ${ieNo} açıldı! ${model} — ${bant} — ${adet} adet`;
}

// ── İŞ EMRİ GÜNCELLE SİDEBARI ────────────────────────────────
function isEmriGuncelleSidebar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ieNolar = ss.getSheetByName("⚙️ PRODUCTION ORDERS").getRange("A4:A300").getValues()
    .filter(r => r[0]).map(r => `<option>${r[0]}</option>`).join("");
  const durumlar = IE_STATUS.map(d => `<option>${d}</option>`).join("");

  _sidebarAc("⚙️ Update ARP", `
<div class="form-group">
  <label>ARP No *</label>
  <select id="ieNo"><option value="">— Seçin —</option>${ieNolar}</select>
</div>
<div class="form-group">
  <label>Yeni Durum *</label>
  <select id="durum"><option value="">— Seçin —</option>${durumlar}</select>
</div>
<div class="form-group">
  <label>Yapılan Adet (toplam)</label>
  <input type="number" id="yapilan" min="0" placeholder="Şimdiye kadar yapılan toplam">
</div>
<div class="form-group" id="nedenDiv" style="display:none">
  <label>Durdurma Nedeni</label>
  <textarea id="neden" placeholder="Why was it stopped?"></textarea>
</div>
<div class="form-group">
  <label>Not</label>
  <textarea id="not" placeholder="Additional notes..."></textarea>
</div>
<button class="btn btn-primary" onclick="kaydet()">✅ Güncelle</button>
<button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
<div id="msg" class="msg"></div>
<script>
document.getElementById('durum').addEventListener('change', function() {
  document.getElementById('nedenDiv').style.display =
    this.value === 'Yarım Bırakıldı' ? 'block' : 'none';
});
function kaydet() {
  const ieNo = document.getElementById('ieNo').value;
  const durum = document.getElementById('durum').value;
  if (!ieNo || !durum) { goster('ARP ve durum seçin!', false); return; }
  document.querySelector('.btn-primary').disabled = true;
  google.script.run
    .withSuccessHandler(r => goster(r, true))
    .withFailureHandler(e => goster(e.message, false))
    .isEmriGuncelle(ieNo, durum,
      parseInt(document.getElementById('yapilan').value) || 0,
      document.getElementById('neden').value,
      document.getElementById('not').value);
}
function goster(m, ok) {
  const d = document.getElementById('msg');
  d.className = 'msg ' + (ok ? 'ok' : 'err');
  d.textContent = m;
  if (ok) setTimeout(() => google.script.host.close(), 2000);
  else { document.querySelector('.btn-primary').disabled = false; }
}
</script>`);
}

function isEmriGuncelle(ieNo, durum, yapilan, neden, not) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName("⚙️ PRODUCTION ORDERS");
  const veriler = s.getRange("A4:A300").getValues();
  for (let i = 0; i < veriler.length; i++) {
    if (veriler[i][0].toString().trim() === ieNo.trim()) {
      const satir = i + 4;
      s.getRange(satir, 8).setValue(durum);
      if (yapilan > 0) s.getRange(satir, 10).setValue(yapilan);
      if (durum === "Completed") s.getRange(satir, 7).setValue(new Date());
      if (neden) s.getRange(satir, 12).setValue(neden);
      if (not) s.getRange(satir, 13).setValue(not);
      s.getRange(satir, 14).setValue(new Date());

      // Sipariş güncelle
      const model = s.getRange(satir, 3).getValue();
      const sipNo = s.getRange(satir, 2).getValue();
      if (durum === "Completed" && model) {
        const sipS = ss.getSheetByName("📋 ORDERS");
        const sipVeri = sipS.getRange("C4:G500").getValues();
        for (let j = 0; j < sipVeri.length; j++) {
          if (sipVeri[j][0].toString().toUpperCase() === model.toUpperCase()) {
            sipS.getRange(j + 4, 5).setValue(yapilan);
            const sipAdet = Number(sipS.getRange(j + 4, 4).getValue()) || 0;
            sipS.getRange(j + 4, 6).setValue(Math.max(0, sipAdet - yapilan));
            if (yapilan >= sipAdet) sipS.getRange(j + 4, 7).setValue("Completed");
            sipS.getRange(j + 4, 12).setValue(new Date());
            break;
          }
        }
      }
      logYaz("İŞ EMRİ", "GÜNCELLEME", `${ieNo} → ${durum} | Yapılan: ${yapilan}`);
      return `✅ ${ieNo} güncellendi: ${durum}`;
    }
  }
  return "⚠️ ARP bulunamadı.";
}

// ── BANT GİRİŞİ SİDEBARI ─────────────────────────────────────
function bantGirisiSidebar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const modeller = ss.getSheetByName("⚙️ SETTINGS").getRange("A2:A100").getValues()
    .filter(r => r[0]).map(r => `<option>${r[0]}</option>`).join("");
  const bantlar = LINES.map(b => `<option>${b}</option>`).join("");
  const ieNolar = ss.getSheetByName("⚙️ PRODUCTION ORDERS").getRange("A4:A300").getValues()
    .filter(r => r[0]).map(r => `<option>${r[0]}</option>`).join("");

  _sidebarAc("📈 Line Entry", `
<div class="form-group">
  <label>Date</label>
  <input type="date" id="tarih" value="${new Date().toISOString().split('T')[0]}">
</div>
<div class="form-group">
  <label>Line *</label>
  <select id="hat"><option value="">— Seçin —</option>${bantlar}</select>
</div>
<div class="form-group">
  <label>Vehicle Model *</label>
  <select id="model"><option value="">— Seçin —</option>${modeller}</select>
</div>
<div class="form-group">
  <label>Color</label>
  <input type="text" id="renk" placeholder="Color info if applicable">
</div>
<div class="form-group">
  <label>Produced Quantity *</label>
  <input type="number" id="adet" min="0" placeholder="Produced today">
</div>
<div class="form-group">
  <label>Total Quantity (cumulative)</label>
  <input type="number" id="toplam" min="0" placeholder="Total produced so far">
</div>
<div class="form-group">
  <label>Açıklama</label>
  <select id="aciklama">
    <option>End of day report</option>
    <option>Bitti</option>
    <option>Paused</option>
  </select>
</div>
<div class="form-group">
  <label>ARP No</label>
  <select id="ieNo"><option value="">— Opsiyonel —</option>${ieNolar}</select>
</div>
<button class="btn btn-primary" onclick="kaydet()">✅ Kaydet</button>
<button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
<div id="msg" class="msg"></div>
<script>
function kaydet() {
  const hat = document.getElementById('hat').value;
  const model = document.getElementById('model').value;
  const adet = document.getElementById('adet').value;
  if (!hat) { goster('Hat seçin!', false); return; }
  if (!model) { goster('Model seçin!', false); return; }
  if (adet === '') { goster('Adet girin!', false); return; }
  document.querySelector('.btn-primary').disabled = true;
  google.script.run
    .withSuccessHandler(r => goster(r, true))
    .withFailureHandler(e => goster(e.message, false))
    .bantKaydet(
      document.getElementById('tarih').value, hat, model,
      document.getElementById('renk').value,
      parseInt(adet)||0,
      parseInt(document.getElementById('toplam').value)||0,
      document.getElementById('aciklama').value,
      document.getElementById('ieNo').value
    );
}
function goster(m, ok) {
  const d = document.getElementById('msg');
  d.className = 'msg ' + (ok ? 'ok' : 'err');
  d.textContent = m;
  if (ok) setTimeout(() => google.script.host.close(), 2000);
  else { document.querySelector('.btn-primary').disabled = false; }
}
</script>`);
}

function bantKaydet(tarih, hat, model, renk, adet, toplam, aciklama, ieNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName("📈 LINE TRACKING");
  const veriler = s.getRange("A4:A500").getValues();
  let sonSatir = 4;
  for (let i = veriler.length - 1; i >= 0; i--) {
    if (veriler[i][0] !== "") { sonSatir = i + 5; break; }
  }
  s.getRange(sonSatir, 1).setValue(tarih ? new Date(tarih) : new Date());
  s.getRange(sonSatir, 2).setValue(hat);
  s.getRange(sonSatir, 3).setValue(model);
  s.getRange(sonSatir, 4).setValue(renk || "");
  s.getRange(sonSatir, 5).setValue(adet);
  s.getRange(sonSatir, 6).setValue(toplam || "");
  s.getRange(sonSatir, 7).setValue(aciklama);
  s.getRange(sonSatir, 8).setValue(ieNo || "");
  logYaz("BANT", "GİRİŞ", `${hat} | ${model} | ${adet} adet | ${aciklama}`);
  return `✅ Saved! ${hat} — ${model} — ${adet} adet`;
}

// ── DASHBOARD YENİLE ─────────────────────────────────────────
function dashboardYenile() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("📊 DASHBOARD");

  // KPI kartlarını güncelle
  const kpi = _dashKpiHesapla(ss);
  const kpiMap = [
    ["A4:B4", kpi.toplamSip],
    ["C4:D4", kpi.tamamlanan],
    ["E4:F4", kpi.uretimde],
    ["G4:H4", kpi.bekliyor],
    ["I4:J4", kpi.saseBekleyen],
    ["K4:L4", kpi.ilerleme + "%"],
  ];
  kpiMap.forEach(([aralik, deger]) => {
    dash.getRange(aralik).setValue(deger);
  });
  dash.getRange("B2:D2").setValue(new Date()).setNumberFormat("dd.mm.yyyy hh:mm");

  // Bant tablosu aşağıda yenileniyor

  // Bant tablosunu yenile — _kurDashboard ile aynı mantık
  _kurDashboardBantTablosu(ss, dash);

  // Sipariş tablosunu temizle ve doldur
  const baslangic = 16;
  dash.getRange(`A${baslangic}:L300`).clearContent().clearFormat();

  const sipData = ss.getSheetByName("📋 ORDERS").getRange("A4:L500").getValues();
  const satirlar = [];
  const renkler = [];

  sipData.forEach(row => {
    if (!row[0] || !row[2]) return;
    const adet    = Number(row[3]) || 0;
    const uretilen = Number(row[4]) || 0;
    const kalan   = Number(row[5]) || 0;
    const durum   = row[6].toString();
    const ilerleme = adet > 0 ? Math.round((uretilen / adet) * 100) : 0;
    const cubuk   = _ilerlermeCubugu(ilerleme);

    satirlar.push([row[2], row[0], adet, uretilen, kalan, durum, cubuk, "", "", "", row[7], row[9]]);

    let renk = "#ffffff";
    if (durum === "Completed") renk = "#d4edda";
    else if (durum === "In Production") renk = "#d1ecf1";
    else if (durum === "Cancelled")    renk = "#f8d7da";
    else if (durum === "Waiting") renk = "#fff3cd";
    renkler.push(renk);
  });

  if (satirlar.length > 0) {
    // Önce merge'leri kaldır
    try { dash.getRange(baslangic, 1, satirlar.length, 12).breakApart(); } catch(e) {}
    SpreadsheetApp.flush();

    dash.getRange(baslangic, 1, satirlar.length, 12).setValues(satirlar);
    dash.getRange(baslangic, 3, satirlar.length, 3).setNumberFormat("#,##0");
    dash.getRange(baslangic, 1, satirlar.length, 1).setWrap(false);

    // Progress bar sütunlarını birleştir ve renk ver
    for (let i = 0; i < satirlar.length; i++) {
      const r = baslangic + i;
      try { dash.getRange(`G${r}:J${r}`).merge(); } catch(e) {}
      dash.getRange(`A${r}:L${r}`).setBackground(renkler[i]);
      dash.getRange(`G${r}`).setFontFamily("Courier New").setFontSize(9)
       .setHorizontalAlignment("left");
      dash.setRowHeight(r, 22);
    }
  }

  logYaz("DASHBOARD", "YENİLENDİ", satirlar.length + " satır");
  SpreadsheetApp.getUi().alert("✅ Dashboard updated! " + satirlar.length + " sipariş gösteriliyor.");
}

function _ilerlermeCubugu(yuzde) {
  const dolu = Math.round(yuzde / 5);  // 20 karakterlik çubuk
  const bos = 20 - dolu;
  const cubuk = "█".repeat(dolu) + "░".repeat(bos);
  return cubuk + "  %" + yuzde;
}

// ── AYARLARI GÖSTER ──────────────────────────────────────────
function ayarlariGoster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ayarlar = ss.getSheetByName("⚙️ SETTINGS");
  ayarlar.showSheet();
  ss.setActiveSheet(ayarlar);

  // Sidebar ile aç — kullanıcı "Bitti" deyince sayfa gizlenir
  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  body { font-family: Arial, sans-serif; padding: 16px; font-size: 13px; background: #f8f9fa; }
  h3 { background: #1a1a2e; color: white; padding: 10px 14px; margin: -16px -16px 16px; font-size: 13px; }
  .item { background: white; border-left: 4px solid #4285f4; padding: 10px 12px;
          margin-bottom: 10px; border-radius: 0 6px 6px 0; }
  .item b { display: block; margin-bottom: 3px; }
  .item span { color: #555; font-size: 12px; }
  .btn { display: block; width: 100%; padding: 10px; background: #1a1a2e;
         color: white; border: none; border-radius: 6px; font-size: 13px;
         font-weight: bold; cursor: pointer; margin-top: 16px; }
  .btn:hover { background: #16213e; }
</style>
</head>
<body>
<h3>⚙️ SETTINGS</h3>
<div class="item">
  <b>A Sütunu — Model Listesi</b>
  <span>Add new models at the bottom of the list</span>
</div>
<div class="item" style="border-color:#34a853">
  <b>C Sütunu — Tedarikçi Listesi</b>
  <span>Add new suppliers at the bottom of the list</span>
</div>
<div class="item" style="border-color:#fbbc04">
  <b>E Sütunu — Bantlar</b>
  <span>Edit to add/change lines</span>
</div>
<p style="font-size:11px;color:#888;margin-top:12px">
  After making changes, press the button below.
</p>
<button class="btn" onclick="bitti()">✅ Done — Close Page</button>
<script>
function bitti() {
  google.script.run
    .withSuccessHandler(() => google.script.host.close())
    .ayarlariKapat();
}
</script>
</body>
</html>
`).setTitle("⚙️ Ayarlar").setWidth(280);
  SpreadsheetApp.getUi().showSidebar(html);
}

function ayarlariKapat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName("⚙️ SETTINGS").hideSheet();
  ss.setActiveSheet(ss.getSheetByName("📊 DASHBOARD"));
}

// ── LOG ───────────────────────────────────────────────────────

function isEmirleriDuzelt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName("⚙️ PRODUCTION ORDERS");
  const veri = s.getRange("E4:J300").getValues();
  const yazilacak = veri.map(r => {
    const planlanan = Number(r[0]) || 0;
    const yapilan   = Number(r[5]) || 0;
    return [planlanan > 0 ? Math.max(0, planlanan - yapilan) : ""];
  });
  s.getRange("K4:K300").clearContent();
  SpreadsheetApp.flush();
  s.getRange("K4:K300").setValues(yazilacak).setNumberFormat("#,##0");
  logYaz("İŞ EMİRLERİ", "DÜZELT", "K sütunu temizlendi");
  SpreadsheetApp.getUi().alert("✅ Errors cleared!");
}

function logYaz(sayfa, islem, detay) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let log = ss.getSheetByName("📝 LOG");
    if (!log) {
      log = ss.insertSheet("📝 LOG");
      log.getRange(1,1,1,5).setValues([["ZAMAN","KULLANICI","SAYFA","İŞLEM","DETAY"]]);
      log.getRange(1,1,1,5).setFontWeight("bold").setBackground(RENKLER.KOYU).setFontColor(RENKLER.BEYAZ);
      log.setFrozenRows(1);
      log.hideSheet();
    }
    log.appendRow([
      new Date(),
      Session.getActiveUser().getEmail() || "-",
      sayfa, islem, detay
    ]);
  } catch(e) {}
}
