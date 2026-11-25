// Google Sheet -> Extensions -> Apps Script -> Paste this code -> Deploy -> New deployment -> Web app (Execute as: Me, Who has access: Anyone)

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("সকল ব্যাচ");
  const data = sheet.getDataRange().getValues();
  const overrides = {};

  // Columns in "সকল ব্যাচ": [0]=ব্যাচ, [1]=শিক্ষাবর্ষ, [2]=ক্রমিক নং, [3]=নাম, [4]=রোল, [5]=রেজি, [6]=মোবাইল, [7]=নোট
  for (let i = 1; i < data.length; i++) {
    const batch = String(data[i][0] || "").trim();
    const serial = String(data[i][2] || "").trim();
    const roll = String(data[i][4] || "").trim();
    const phone = String(data[i][6] || "").trim();
    if (batch && serial) {
      overrides[batch + "|" + serial + "|" + roll] = phone;
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, phones: overrides }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const batch = String(body.batch || "").trim();
    const serial = String(body.serial || "").trim();
    const roll = String(body.roll || "").trim();
    const phone = String(body.phone || "").trim();
    const formattedPhone = phone && phone.startsWith("0") ? "'" + phone : phone;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let updated = 0;

    // 1. Update in "সকল ব্যাচ" (col G = 7)
    const master = ss.getSheetByName("সকল ব্যাচ");
    if (master) {
      const mData = master.getDataRange().getValues();
      for (let i = 1; i < mData.length; i++) {
        if (
          String(mData[i][0] || "").trim() === batch &&
          String(mData[i][2] || "").trim() === serial &&
          String(mData[i][4] || "").trim() === roll
        ) {
          master.getRange(i + 1, 7).setValue(formattedPhone);
          updated++;
          break;
        }
      }
    }

    // 2. Update in individual batch tab (col E = 5)
    const tab = ss.getSheetByName(batch);
    if (tab) {
      const tData = tab.getDataRange().getValues();
      for (let i = 1; i < tData.length; i++) {
        if (
          String(tData[i][0] || "").trim() === serial &&
          String(tData[i][2] || "").trim() === roll
        ) {
          tab.getRange(i + 1, 5).setValue(formattedPhone);
          updated++;
          break;
        }
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, updated }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
