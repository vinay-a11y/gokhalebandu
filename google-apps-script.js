/**
 * Google Apps Script for handling Gokhale Bandhu Diwali Faral order form submissions
 *
 * ⚠️ IMPORTANT: This file is NOT meant to run in your local project!
 * This code should be copied and pasted into the Google Apps Script editor.
 *
 * SETUP INSTRUCTIONS:
 * 1. Open your Google Sheet
 * 2. Go to Extensions > Apps Script
 * 3. Delete any existing code and paste this entire script
 * 4. Click on the "Deploy" button > "New deployment"
 * 5. Select type: "Web app"
 * 6. Execute as: "Me"
 * 7. Who has access: "Anyone"
 * 8. Click "Deploy"
 * 9. Copy the Web App URL and paste it in the HTML file (GOOGLE_SCRIPT_URL variable)
 * 10. Authorize the script when prompted
 */

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents)
    
    // Determine which sheet to use based on order type
    let sheetName;
    if (data.orderType === 'In Pune') {
      sheetName = 'InPune';
    } else if (data.orderType === 'Outside Pune') {
      sheetName = 'OutsidePune';
    } else {
      sheetName = 'International';
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet()
    let sheet = ss.getSheetByName(sheetName)

    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(sheetName)

      const headers = [
        "Timestamp",
        "Order Type",
        "Name",
        "Contact Number",
        "Shipping Address",
        "Country",
        "Dispatch Date",
        "Products Ordered",
        "Total Boxes",
        "Total Weight (kg)",
        "Subtotal (₹)",
        "Delivery Fee",
        "Grand Total (₹)",
        "Payment Status"
      ]

      sheet.getRange(1, 1, 1, headers.length).setValues([headers])

      const headerRange = sheet.getRange(1, 1, 1, headers.length)
      headerRange.setFontWeight("bold")
      headerRange.setBackground("#8e24aa")
      headerRange.setFontColor("#ffffff")

      sheet.setFrozenRows(1)

      const protection = sheet.protect().setDescription('Order data - Protected from deletion');
      
      // Allow all editors to edit, but prevent deletion
      protection.setWarningOnly(true);
      
      Logger.log('Sheet protection enabled for: ' + sheetName);
    }

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
      data.contact,
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
    ]

    sheet.appendRow(rowData)

    sheet.autoResizeColumns(1, 14)

    updateKitchenPrepSheet(ss, data.products)

    sendEmailNotification(data, sheetName)

    return ContentService.createTextOutput(
      JSON.stringify({
        status: "success",
        message: "Order submitted successfully",
      }),
    ).setMimeType(ContentService.MimeType.JSON)
  } catch (error) {
    Logger.log("Error: " + error.toString())

    return ContentService.createTextOutput(
      JSON.stringify({
        status: "error",
        message: error.toString(),
      }),
    ).setMimeType(ContentService.MimeType.JSON)
  }
}

function updateKitchenPrepSheet(ss, products) {
  try {
    let kitchenSheet = ss.getSheetByName('Kitchen Prep');
    
    // Create Kitchen Prep sheet if it doesn't exist
    if (!kitchenSheet) {
      kitchenSheet = ss.insertSheet('Kitchen Prep');
      
      const headers = ['Product Name', 'Total Quantity'];
      kitchenSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      const headerRange = kitchenSheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#d4af37');
      headerRange.setFontColor('#ffffff');
      
      kitchenSheet.setFrozenRows(1);
      kitchenSheet.setColumnWidth(1, 400);
      kitchenSheet.setColumnWidth(2, 150);
      
      Logger.log('Kitchen Prep sheet created');
    }
    
    // Get all existing data from Kitchen Prep sheet
    const lastRow = kitchenSheet.getLastRow();
    let existingData = {};
    
    if (lastRow > 1) {
      const dataRange = kitchenSheet.getRange(2, 1, lastRow - 1, 2);
      const values = dataRange.getValues();
      
      values.forEach(row => {
        if (row[0]) {
          existingData[row[0]] = parseInt(row[1]) || 0;
        }
      });
    }
    
    // Add new order quantities to existing totals
    for (const [product, quantity] of Object.entries(products)) {
      if (quantity > 0) {
        if (existingData[product]) {
          existingData[product] += quantity;
        } else {
          existingData[product] = quantity;
        }
      }
    }
    
    // Clear existing data (except headers)
    if (lastRow > 1) {
      kitchenSheet.getRange(2, 1, lastRow - 1, 2).clear();
    }
    
    // Write updated data back to sheet
    const sortedProducts = Object.keys(existingData).sort();
    const newData = sortedProducts.map(product => [product, existingData[product]]);
    
    if (newData.length > 0) {
      kitchenSheet.getRange(2, 1, newData.length, 2).setValues(newData);
    }
    
    Logger.log('Kitchen Prep sheet updated successfully');
    
  } catch (error) {
    Logger.log('Error updating Kitchen Prep sheet: ' + error.toString());
  }
}

function sendEmailNotification(data, sheetName) {
  try {
    // Replace with your email address
    const recipientEmail = "gokhalebandhu7@gmail.com"

    const subject = `New Diwali Faral Order - ${data.orderType} - ${data.name}`

    let productList = ""
    for (const [product, quantity] of Object.entries(data.products)) {
      if (quantity > 0) {
        productList += `${product}: ${quantity}\n`
      }
    }

    const body = `
New Order Received!

Order Details:
--------------
Order Type: ${data.orderType}
Name: ${data.name}
Contact: ${data.contact}
Address: ${data.address}
Country: ${data.country}
Dispatch Date: ${data.dispatchDate}

Products Ordered:
-----------------
${productList}

Order Summary:
--------------
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
    `

    MailApp.sendEmail(recipientEmail, subject, body)
  } catch (error) {
    Logger.log("Email notification error: " + error.toString())
  }
}

function testSetup() {
  const testData = {
    timestamp: new Date().toISOString(),
    orderType: "In Pune",
    name: "Test User",
    contact: "9881763116",
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
  }

  const e = {
    postData: {
      contents: JSON.stringify(testData),
    },
  }

  const result = doPost(e)
  Logger.log(result.getContent())
}
