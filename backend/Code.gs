/**
 * UCOP SPIKES — Backend (Google Apps Script)
 * ------------------------------------------
 * Bound to a Google Sheet with 2 tabs: "Orders" and "Settings"
 * Deploy this as a Web App (Execute as: Me, Who has access: Anyone)
 * Copy the deployment URL into order.html and admin.html (API_URL constant)
 */

// ===== CONFIG =====
const ADMIN_PASSWORD = "ucop2026"; // TODO: change this to your own admin password
const DRIVE_FOLDER_NAME = "UCOP Spikes - Payment Proofs";

// ===== SHEET HELPERS =====
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === "Orders") {
      sheet.appendRow([
        "Order ID", "Timestamp", "Customer Name", "Phone", "Product", "Colour", "Size",
        "Qty", "Unit Price (RM)", "Total Price (RM)", "Delivery Method", "Address / Notes",
        "Payment Method", "Payment Proof", "Status", "Tracking / Remarks", "Admin Notes"
      ]);
    }
    if (name === "Settings") {
      sheet.appendRow(["Key", "Value"]);
      sheet.appendRow(["WHATSAPP_NUMBER", "601126103958"]);
      sheet.appendRow(["BANK_NAME", "TODO: your bank name"]);
      sheet.appendRow(["BANK_ACCOUNT_NO", "TODO: your account number"]);
      sheet.appendRow(["BANK_ACCOUNT_NAME", "TODO: account holder name"]);
      sheet.appendRow(["QR_IMAGE_URL", "TODO: paste your DuitNow/TnG QR image URL here"]);
    }
  }
  return sheet;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function generateOrderId() {
  const sheet = getSheet("Orders");
  const lastRow = sheet.getLastRow();
  const num = lastRow < 1 ? 1 : lastRow; // header row counts as 0 orders
  const padded = String(num).padStart(5, "0");
  return "UCOP-" + padded;
}

// ===== PAYMENT PROOF UPLOAD =====
function saveProofImage(base64Data, filename, mimeType) {
  let folder;
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);

  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, mimeType, filename);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// ===== doGet: settings / orders (admin) / report =====
function doGet(e) {
  const action = e.parameter.action;

  if (action === "getSettings") {
    const sheet = getSheet("Settings");
    const data = sheet.getDataRange().getValues();
    const settings = {};
    for (let i = 1; i < data.length; i++) {
      settings[data[i][0]] = data[i][1];
    }
    return jsonOut({ success: true, settings: settings });
  }

  if (action === "getOrders") {
    if (e.parameter.password !== ADMIN_PASSWORD) {
      return jsonOut({ success: false, error: "Unauthorized" });
    }
    const sheet = getSheet("Orders");
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const orders = [];
    for (let i = 1; i < data.length; i++) {
      const row = {};
      headers.forEach((h, idx) => { row[h] = data[i][idx]; });
      row["_row"] = i + 1; // actual sheet row number for updates
      orders.push(row);
    }
    return jsonOut({ success: true, orders: orders });
  }

  return jsonOut({ success: false, error: "Unknown action" });
}

// ===== doPost: create order / update status =====
function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const action = body.action;

  if (action === "createOrder") {
    const sheet = getSheet("Orders");
    const orderId = generateOrderId();

    let proofUrl = "";
    if (body.proofBase64 && body.proofFilename) {
      try {
        proofUrl = saveProofImage(body.proofBase64, body.proofFilename, body.proofMimeType || "image/jpeg");
      } catch (err) {
        proofUrl = "Upload failed: " + err.message;
      }
    }

    sheet.appendRow([
      orderId,
      new Date(),
      body.customerName || "",
      body.phone || "",
      body.product || "",
      body.colour || "",
      body.size || "",
      body.qty || 1,
      body.unitPrice || 0,
      body.totalPrice || 0,
      body.deliveryMethod || "",
      body.addressNotes || "",
      body.paymentMethod || "",
      proofUrl,
      "Pending Payment Verification",
      "",
      ""
    ]);

    return jsonOut({ success: true, orderId: orderId });
  }

  if (action === "updateStatus") {
    if (body.password !== ADMIN_PASSWORD) {
      return jsonOut({ success: false, error: "Unauthorized" });
    }
    const sheet = getSheet("Orders");
    const row = body.row;
    const statusColIndex = 15; // "Status" is column O (1-indexed = 15)
    sheet.getRange(row, statusColIndex).setValue(body.status);
    if (body.trackingRemarks !== undefined) {
      sheet.getRange(row, 16).setValue(body.trackingRemarks); // Tracking / Remarks column
    }
    if (body.adminNotes !== undefined) {
      sheet.getRange(row, 17).setValue(body.adminNotes); // Admin Notes column
    }
    return jsonOut({ success: true });
  }

  return jsonOut({ success: false, error: "Unknown action" });
}
