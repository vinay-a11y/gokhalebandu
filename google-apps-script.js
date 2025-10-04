/**
 * Google Apps Script for handling Gokhale Bandhu Diwali Faral order form submissions
 * 
 * Features:
 * - Prevents other editors from deleting sheets (only owner can).
 * - Automatic daily backup into a separate spreadsheet.
 * - Handles orders, updates kitchen prep, sends email notifications.
 * - Sends order confirmation to customer email
 */

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    
    // Determine sheet based on order type
    let sheetName;
    if (data.orderType === 'In Pune') {
      sheetName = 'InPune';
    } else if (data.orderType === 'Outside Pune') {
      sheetName = 'OutsidePune';
    } else {
      sheetName = 'International';
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);

    // Create sheet if not exists
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);

      const headers = [
        "Timestamp", "Order Type", "Name", "Contact Number", "Shipping Address",
        "Country", "Dispatch Date", "Products Ordered", "Total Boxes",
        "Total Weight (kg)", "Subtotal (₹)", "Delivery Fee",
        "Grand Total (₹)", "Payment Status"
      ];

      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#8e24aa").setFontColor("#ffffff");
      sheet.setFrozenRows(1);

      protectSheet(sheet);
    }

    // Format ordered products
    let productsOrdered = '';
    for (const [product, quantity] of Object.entries(data.products)) {
      if (quantity > 0) {
        productsOrdered += `${product}: ${quantity}\n`;
      }
    }

    const rowData = [
      new Date(data.timestamp),
      data.orderType,
      data.name,
      data.contact, // Stores "phone/email" format
      data.address,
      data.country,
      data.dispatchDate,
      productsOrdered,
      data.totalBoxes || 0,
      data.totalWeightKg || 0,
      data.subtotal || 0,
      data.deliveryFee || 'To be informed',
      data.grandTotal || 0,
      'Payment Confirmed'
    ];

    sheet.appendRow(rowData);
    sheet.autoResizeColumns(1, 14);

    updateKitchenPrepSheet(ss, data.products);
    sendEmailNotification(data, sheetName);
    sendCustomerConfirmationEmail(data);

    return ContentService.createTextOutput(
      JSON.stringify({ status: "success", message: "Order submitted successfully" })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log("Error: " + error.toString());
    return ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: error.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Protect a sheet so only the owner can edit/delete it
 */
function protectSheet(sheet) {
  const protection = sheet.protect().setDescription("Protected from deletion");
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
  Logger.log("Protection applied to sheet: " + sheet.getName());
}

/**
 * Update Kitchen Prep sheet with total quantities
 */
function updateKitchenPrepSheet(ss, products) {
  try {
    let kitchenSheet = ss.getSheetByName('Kitchen Prep');
    if (!kitchenSheet) {
      kitchenSheet = ss.insertSheet('Kitchen Prep');
      const headers = ['Product Name', 'Total Quantity'];
      kitchenSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      kitchenSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d4af37').setFontColor('#ffffff');
      kitchenSheet.setFrozenRows(1);
      kitchenSheet.setColumnWidth(1, 400);
      kitchenSheet.setColumnWidth(2, 150);
      protectSheet(kitchenSheet);
    }

    const lastRow = kitchenSheet.getLastRow();
    let existingData = {};

    if (lastRow > 1) {
      const values = kitchenSheet.getRange(2, 1, lastRow - 1, 2).getValues();
      values.forEach(row => {
        if (row[0]) {
          existingData[row[0]] = parseInt(row[1]) || 0;
        }
      });
    }

    for (const [product, quantity] of Object.entries(products)) {
      if (quantity > 0) {
        existingData[product] = (existingData[product] || 0) + quantity;
      }
    }

    if (lastRow > 1) {
      kitchenSheet.getRange(2, 1, lastRow - 1, 2).clear();
    }

    const sortedProducts = Object.keys(existingData).sort();
    const newData = sortedProducts.map(product => [product, existingData[product]]);
    if (newData.length > 0) {
      kitchenSheet.getRange(2, 1, newData.length, 2).setValues(newData);
    }
  } catch (error) {
    Logger.log('Error updating Kitchen Prep sheet: ' + error.toString());
  }
}

/**
 * Send email notification to admin on new order
 */
function sendEmailNotification(data, sheetName) {
  try {
    const recipientEmail = "gokhalebandhu7@gmail.com";
    const subject = `New Diwali Faral Order - ${data.orderType} - ${data.name}`;

    let productList = "";
    for (const [product, quantity] of Object.entries(data.products)) {
      if (quantity > 0) {
        productList += `${product}: ${quantity}\n`;
      }
    }

    // Extract phone and email from contact field
    const contactParts = data.contact.split('/');
    const phone = contactParts[0] || data.contact;
    const email = contactParts[1] || 'Not provided';

    const body = `
New Order Received!

Order Type: ${data.orderType}
Name: ${data.name}
Contact: ${phone}
Email: ${email}
Address: ${data.address}
Country: ${data.country}
Dispatch Date: ${data.dispatchDate}

Products Ordered:
${productList}

Order Summary:
Total Boxes: ${data.totalBoxes || 0}
Total Weight: ${data.totalWeightKg || 0} kg
Subtotal: ₹${data.subtotal || 0}
Delivery Fee: ${data.deliveryFee || 'To be informed'}
Grand Total: ₹${data.grandTotal || 0}
Payment Status: Payment Confirmed

Sheet: ${sheetName}
Timestamp: ${new Date(data.timestamp).toLocaleString()}

---
Gokhale Bandhu - Diwali Faral Orders
Contact: +91 9881763116
    `;
    MailApp.sendEmail(recipientEmail, subject, body);
  } catch (error) {
    Logger.log("Email notification error: " + error.toString());
  }
}

/**
 * Send order confirmation email to customer
 */
function sendCustomerConfirmationEmail(data) {
  try {
    // Extract email from contact field (format: "phone/email")
    const contactParts = data.contact.split('/');
    if (contactParts.length < 2) {
      Logger.log("No email found in contact field");
      return;
    }
    
    const customerEmail = contactParts[1].trim();
    const phone = contactParts[0].trim();
    
    const subject = `Order Confirmation - Gokhale Bandhu Diwali Faral - ${data.name}`;

    let productList = "";
    for (const [product, quantity] of Object.entries(data.products)) {
      if (quantity > 0) {
        productList += `  • ${product}: ${quantity} box(es)\n`;
      }
    }

    const body = `
Dear ${data.name},

Thank you for your order! 🪔

We are delighted to confirm your Diwali Faral order with Gokhale Bandhu.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ORDER DETAILS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Order Type: ${data.orderType}
Dispatch Date: ${data.dispatchDate}
Order Date: ${new Date(data.timestamp).toLocaleString()}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PRODUCTS ORDERED
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

${productList}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ORDER SUMMARY
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Total Boxes: ${data.totalBoxes || 0}
Total Weight: ${data.totalWeightKg || 0} kg
Subtotal: ₹${data.subtotal || 0}
Delivery Fee: ${data.deliveryFee || 'To be informed'}
Grand Total: ₹${data.grandTotal || 0}

Payment Status: ✅ Payment Confirmed

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
DELIVERY INFORMATION
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Shipping Address:
${data.address}
${data.country !== 'N/A' ? 'Country: ' + data.country : ''}

Dispatched From:
Shed No. 1, Behind Ajinkya Nagar Society, Nityanand Hall Lane,
Hinge Sinhgad, Pune, Maharashtra, India

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STORAGE INSTRUCTIONS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

• Motichoor Ladoo: Best consumed within 8-10 days. 
  For longer shelf life, store in the refrigerator.

• All other products: Store in an airtight container; 
  best within 2 months.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
CONTACT US
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

For any questions or concerns about your order:

📞 Phone: +91 9881763116
✉️ Email: custcare.virtue@gmail.com
📷 Instagram: @gokhale_bandhu
📘 Facebook: Gokhale Bandhu

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

✨ Brand: Gokhale Bandhu ✨
🥥 Uses Groundnut Refined Oil | 🧈 Uses Pure Ghee
🚫 No Colours & Preservatives

Thank you for choosing Gokhale Bandhu!
Wishing you a Happy and Prosperous Diwali! 🪔

Warm regards,
Gokhale Bandhu Team
    `;

    MailApp.sendEmail(customerEmail, subject, body);
    Logger.log("Customer confirmation email sent to: " + customerEmail);
  } catch (error) {
    Logger.log("Customer confirmation email error: " + error.toString());
  }
}

/**
 * Daily Backup - Copies all sheets into another spreadsheet
 */
function backupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backupFolderId = "1m4xfXB-PYTcS-XJ_Rt7Zu1MVWEPflu_E"; // <-- create a folder in Drive and paste its ID here

  const date = new Date();
  const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm");

  const backupFileName = ss.getName() + "_Backup_" + formattedDate;
  const backup = ss.copy(backupFileName);

  DriveApp.getFileById(backup.getId()).moveTo(DriveApp.getFolderById(backupFolderId));

  Logger.log("Backup created: " + backupFileName);
}

/**
 * Setup daily trigger for backup
 */
function createDailyBackupTrigger() {
  ScriptApp.newTrigger("backupSheets")
    .timeBased()
    .everyDays(1)
    .atHour(0) // runs at midnight
    .create();
}

/**
 * Test Setup
 */
function testSetup() {
  const testData = {
    timestamp: new Date().toISOString(),
    orderType: "In Pune",
    name: "Test User",
    contact: "9881763116/test@example.com",
    address: "Test Address, Pune",
    country: "N/A",
    dispatchDate: "13/10/2025",
    products: {
      "Bhajani Chakali (200gm)": 2,
      "Bhajani Chakali (500gm)": 1,
      "Besan Ladoo (200gm)": 1,
      "Motichoor Ladoo (500gm)": 1
    },
    totalBoxes: 5,
    totalWeightKg: 1.5,
    subtotal: 925,
    deliveryFee: "₹80",
    grandTotal: 1005,
  };

  const e = { postData: { contents: JSON.stringify(testData) } };
  const result = doPost(e);
  Logger.log(result.getContent());
}
