// ============================================================
// Code.gs — Google Apps Script REST API
// คลังบันทึกบิลใบเสร็จ
// ============================================================

// ===== ตั้งค่า =====
const SHEET_NAME = "Receipts";
const DRIVE_FOLDER_NAME = "ReceiptImages"; // ชื่อโฟลเดอร์ใน Google Drive
const CACHE_KEY = "receipts_cache";
const CACHE_EXPIRY = 300; // วินาที (5 นาที)

// ===== CORS Headers =====
function setCORSHeaders(output) {
  return output
    .setHeader("Access-Control-Allow-Origin", "*")
    .setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
    .setHeader("Access-Control-Allow-Headers", "Content-Type");
}

// ===== doGet =====
function doGet(e) {
  try {
    const action = e.parameter.action || "list";

    if (action === "list") {
      return handleList();
    }

    return jsonResponse({ success: false, error: "Unknown action" }, 400);
  } catch (err) {
    return jsonResponse({ success: false, error: err.message }, 500);
  }
}

// ===== doPost =====
function doPost(e) {
  try {
    // OPTIONS preflight
    if (!e.postData) {
      return jsonResponse({ success: true, message: "OK" });
    }

    const body = JSON.parse(e.postData.contents);
    const action = body.action || "save";

    if (action === "save") {
      return handleSave(body);
    }

    if (action === "delete") {
      return handleDelete(body);
    }

    return jsonResponse({ success: false, error: "Unknown action" }, 400);
  } catch (err) {
    return jsonResponse({ success: false, error: err.message }, 500);
  }
}

// ===== Handle List =====
function handleList() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY);

  if (cached) {
    const output = ContentService.createTextOutput(cached)
      .setMimeType(ContentService.MimeType.JSON);
    return setCORSHeaders(output);
  }

  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return jsonResponse({ success: true, data: [] });
  }

  const headers = data[0]; // Date, Store, Description, Amount, ImageURL, Status, Timestamp
  const rows = data.slice(1);

  const receipts = rows
    .filter(row => row[5] !== "ซ่อน") // กรอง Status ≠ 'ซ่อน'
    .map((row, index) => ({
      id: index + 2, // row number (1-indexed, +1 for header)
      date: formatDate(row[0]),
      store: row[1] || "",
      description: row[2] || "",
      amount: parseFloat(row[3]) || 0,
      imageURL: row[4] || "",
      status: row[5] || "",
      timestamp: row[6] ? new Date(row[6]).toISOString() : ""
    }))
    .sort((a, b) => new Date(b.date) - new Date(a.date)); // เรียงใหม่ → เก่า

  const result = JSON.stringify({ success: true, data: receipts });
  cache.put(CACHE_KEY, result, CACHE_EXPIRY);

  return jsonResponse({ success: true, data: receipts });
}

// ===== Handle Save =====
function handleSave(body) {
  // Validate required fields
  if (!body.date) return jsonResponse({ success: false, error: "กรุณาระบุวันที่" }, 400);
  if (!body.store) return jsonResponse({ success: false, error: "กรุณาระบุร้านค้า" }, 400);
  if (!body.amount && body.amount !== 0) return jsonResponse({ success: false, error: "กรุณาระบุยอดเงิน" }, 400);

  let imageURL = "";

  // อัปโหลดรูปภาพถ้ามี
  if (body.imageBase64 && body.imageBase64.length > 0) {
    try {
      imageURL = uploadImageToDrive(body.imageBase64, body.imageName || "receipt.jpg", body.imageMimeType || "image/jpeg");
    } catch (imgErr) {
      // ไม่ block การบันทึก แค่ log error
      console.error("Image upload failed: " + imgErr.message);
    }
  }

  const sheet = getSheet();
  const timestamp = new Date();

  sheet.appendRow([
    body.date,
    body.store,
    body.description || "",
    parseFloat(body.amount) || 0,
    imageURL,
    body.status || "ปกติ",
    timestamp
  ]);

  // Clear cache
  CacheService.getScriptCache().remove(CACHE_KEY);

  return jsonResponse({ success: true, message: "บันทึกสำเร็จ", imageURL: imageURL });
}

// ===== Handle Delete (ซ่อนแถว) =====
function handleDelete(body) {
  if (!body.id) return jsonResponse({ success: false, error: "ไม่พบ ID" }, 400);

  const sheet = getSheet();
  const rowIndex = parseInt(body.id);

  if (rowIndex < 2) return jsonResponse({ success: false, error: "ID ไม่ถูกต้อง" }, 400);

  sheet.getRange(rowIndex, 6).setValue("ซ่อน"); // เซต Status = 'ซ่อน'

  CacheService.getScriptCache().remove(CACHE_KEY);

  return jsonResponse({ success: true, message: "ลบสำเร็จ" });
}

// ===== อัปโหลดรูปไป Google Drive =====
function uploadImageToDrive(base64Data, fileName, mimeType) {
  // ตัด data URL prefix ออก (ถ้ามี)
  const base64Clean = base64Data.replace(/^data:[^;]+;base64,/, "");

  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64Clean),
    mimeType,
    fileName
  );

  const folder = getDriveFolder();
  const file = folder.createFile(blob);

  // ตั้งค่าให้ทุกคนดูได้
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // แปลงเป็น Direct Image Link
  const fileId = file.getId();
  return `https://drive.google.com/uc?export=view&id=${fileId}`;
}

// ===== Helper: Get or Create Folder =====
function getDriveFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

// ===== Helper: Get or Create Sheet =====
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // สร้าง Header
    sheet.appendRow(["Date", "Store", "Description", "Amount", "ImageURL", "Status", "Timestamp"]);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold");
  }

  return sheet;
}

// ===== Helper: Format Date =====
function formatDate(value) {
  if (!value) return "";
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(value);
}

// ===== Helper: JSON Response =====
function jsonResponse(obj, statusCode) {
  const output = ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return setCORSHeaders(output);
}
