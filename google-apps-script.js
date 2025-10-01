/**
 * Google Apps Script for handling form submissions
 *
 * ⚠️ IMPORTANT: This file is NOT meant to run in your local project!
 * This code should be copied and pasted into the Google Apps Script editor.
 *
 * SpreadsheetApp, ContentService, Logger, and MailApp are built-in global
 * objects provided by Google Apps Script environment.
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

// Configuration
const SHEET_NAME = "Orders" // Change this to your sheet name

function doPost(e) {
  try {
    // Parse the incoming JSON data
    const data = JSON.parse(e.postData.contents)

    // Get the active spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    let sheet = ss.getSheetByName(SHEET_NAME)

    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME)

      const headers = [
        "Timestamp",
        "Order Type",
        "Email",
        "Name",
        "Contact Number",
        "Shipping Address",
        "Country",
        "Dispatch Date",
        "Saffron Flavoured Motichoor Ladoo",
        "Besan Ladoo",
        "Paushtik Ladoo",
        "Sweet Whole Wheat Shankarpale with Jaggery",
        "Khare Shankarpale",
        "Bhajani Chakali",
        "Bhajani Kadboli",
        "Plain Shev",
        "Garlic Shev",
        "Premium Chivada",
        "Anarase",
        "Dink Ladoo",
        "Nachani-Moog Ladoo",
        "Ola Naral Karanji",
        "Puranpoli",
        "Diwali Faral Gift Box 1",
        "Diwali Faral Gift Box 2",
        "Total Boxes",
        "Subtotal (₹)",
        "Delivery Fee (₹)",
        "Grand Total (₹)",
      ]

      sheet.getRange(1, 1, 1, headers.length).setValues([headers])

      const headerRange = sheet.getRange(1, 1, 1, headers.length)
      headerRange.setFontWeight("bold")
      headerRange.setBackground("#8e24aa")
      headerRange.setFontColor("#ffffff")

      // Freeze header row
      sheet.setFrozenRows(1)
    }

    // Prepare row data
    const productNames = [
      "Saffron Flavoured Motichoor Ladoo",
      "Besan Ladoo",
      "Paushtik Ladoo",
      "Sweet Whole Wheat Shankarpale with Jaggery",
      "Khare Shankarpale",
      "Bhajani Chakali",
      "Bhajani Kadboli",
      "Plain Shev",
      "Garlic Shev",
      "Premium Chivada",
      "Anarase",
      "Dink Ladoo",
      "Nachani-Moog Ladoo",
      "Ola Naral Karanji",
      "Puranpoli",
      "Diwali Faral Gift Box 1",
      "Diwali Faral Gift Box 2",
    ]

    let totalBoxes = 0
    const productQuantities = productNames.map((name) => {
      const qty = data.products[name] || 0
      totalBoxes += qty
      return qty
    })

    const rowData = [
      new Date(data.timestamp),
      data.orderType,
      data.email,
      data.name,
      data.contact,
      data.address,
      data.country,
      data.dispatchDate,
      ...productQuantities,
      data.totalBoxes || totalBoxes,
      data.subtotal || 0,
      data.deliveryFee || 0,
      data.grandTotal || 0,
    ]

    // Append the data to the sheet
    sheet.appendRow(rowData)

    // Auto-resize columns for better readability
    sheet.autoResizeColumns(1, rowData.length)

    sendEmailNotification(data, totalBoxes)

    // Return success response
    return ContentService.createTextOutput(
      JSON.stringify({
        status: "success",
        message: "Order submitted successfully",
      }),
    ).setMimeType(ContentService.MimeType.JSON)
  } catch (error) {
    // Log error
    Logger.log("Error: " + error.toString())

    // Return error response
    return ContentService.createTextOutput(
      JSON.stringify({
        status: "error",
        message: error.toString(),
      }),
    ).setMimeType(ContentService.MimeType.JSON)
  }
}

function sendEmailNotification(data, totalBoxes) {
  try {
    // Replace with your email address
    const recipientEmail = "gokhalebandhugokhalebandhu7@gmail.com"

    const subject = `New Diwali Faral Order from ${data.name}`

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
Email: ${data.email}
Contact: ${data.contact}
Address: ${data.address}
Country: ${data.country}
Dispatch Date: ${data.dispatchDate}

Products Ordered:
-----------------
${productList}

Order Summary:
--------------
Total Boxes: ${data.totalBoxes || totalBoxes}
Subtotal: ₹${data.subtotal || 0}
Delivery Fee: ${data.deliveryFee === 0 ? 'FREE' : '₹' + (data.deliveryFee || 0)}
Grand Total: ₹${data.grandTotal || 0}

Timestamp: ${new Date(data.timestamp).toLocaleString()}
    `

    MailApp.sendEmail(recipientEmail, subject, body)
  } catch (error) {
    Logger.log("Email notification error: " + error.toString())
    // Don't throw error - email is optional
  }
}

function testSetup() {
  const testData = {
    timestamp: new Date().toISOString(),
    orderType: "Domestic",
    email: "test@example.com",
    name: "Test User",
    contact: "1234567890",
    address: "Test Address",
    country: "N/A",
    dispatchDate: "13/10/2025",
    products: {
      "Saffron Flavoured Motichoor Ladoo": 2,
      "Besan Ladoo": 1,
      "Paushtik Ladoo": 0,
      "Sweet Whole Wheat Shankarpale with Jaggery": 0,
      "Khare Shankarpale": 0,
      "Bhajani Chakali": 0,
      "Bhajani Kadboli": 0,
      "Plain Shev": 0,
      "Garlic Shev": 0,
      "Premium Chivada": 0,
      Anarase: 0,
      "Dink Ladoo": 0,
      "Nachani-Moog Ladoo": 0,
      "Ola Naral Karanji": 0,
      Puranpoli: 0,
      "Diwali Faral Gift Box 1": 1,
      "Diwali Faral Gift Box 2": 0,
    },
    totalBoxes: 4,
    subtotal: 975,
    deliveryFee: 90,
    grandTotal: 1065,
  }

  const e = {
    postData: {
      contents: JSON.stringify(testData),
    },
  }

  const result = doPost(e)
  Logger.log(result.getContent())
}
