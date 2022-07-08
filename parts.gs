const endpoint = "https://data.mongodb-api.com/app/data-rapuj/endpoint/data/v1"
const clusterName = "Cluster1"
const databaseName = "partsDB"
const collectionName = "parts"

const columnVals = ["Part Name", "Barcode", "Datasheet", "Is Consumable", "Description", "Stock", "Category", "Image URL"]

function getAPIKey() {
  const userProperties = PropertiesService.getUserProperties();
  let apikey = userProperties.getProperty('APIKEY');
  let resetKey = false; //Make true if you have to change key
  if (apikey == null || resetKey ) {
    var result = SpreadsheetApp.getUi().prompt(
    'Enter API Key',
    'Key:', SpreadsheetApp.getUi().ButtonSet);
    apikey = result.getResponseText()
    userProperties.setProperty('APIKEY', apikey);
  }
  return apikey;
}

function changeAPIKey() {
  const userProperties = PropertiesService.getUserProperties();
  var result = SpreadsheetApp.getUi().prompt(
  'Enter API Key',
  'Key:', SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() == SpreadsheetApp.getUi().Button.OK) {
    const apikey = result.getResponseText()
    userProperties.setProperty('APIKEY', apikey);
  }
}

function checkFormatting() {
  var sheet = SpreadsheetApp.getActiveSheet()
  var data = sheet.getDataRange().getValues()
  Logger.log(data)
  // loop through each row
  for (var row = 1; row < data.length; row++) {
    for (var col = 0; col < columnVals.length; col++) {
      
      // check for empty cells
      if (data[row][col] === "") {
        SpreadsheetApp.getUi().alert("Empty cell at \nRow: " + (row + 1).toString() + ". \nColumn: " + columnVals[col])
        return false
      }
      
      // make sure isConsumable is set to true/false
      if (col === 3) {
        if (data[row][col] !== true && data[row][col] !== false) {
          SpreadsheetApp.getUi().alert("Issue with:\nRow: " + (row + 1).toString() + ". \nColumn: " + columnVals[col] + "\nIs Consumable column must be set to true or false")
          return false
        }
      }

      // check that stock is an integer value
      if (col === 5) {
        // if parsing the element to an int returns a NaN, then it's probably not an integer
        if (isNaN(parseInt(data[row][col]))) {
          SpreadsheetApp.getUi().alert("Stock must be a number:\nRow: " + (row + 1).toString() + ". \nColumn: " + columnVals[col])
          return false
        }
      }
    }
  }

  SpreadsheetApp.getUi().alert("No formatting issues detected")
  return true
}

function findPart(partBarcode, apikey) {
  // create find endpoint
  const findEndpoint = endpoint + "/action/findOne"
  
  // query to search via barcode
  const query = {barcode: partBarcode}

  const payload = {
    filter: query,
    collection: collectionName,
    database: databaseName,
    dataSource: clusterName
  }

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: { "api-key": apikey }
 }

  const response = UrlFetchApp.fetch(findEndpoint, options);

  // parse response object to a JS Object
  const parsedResponse = JSON.parse(response.getContentText())

  if (parsedResponse.document == null) {
    // part didn't exist in database
    return false
  }
  else {
    // part exists in database
    return true
  }

}

function insertParts() {
  const apikey = getAPIKey()
  // add insert action to endpoint
  const insertEndpoint = endpoint + "/action/insertOne"

  var duplicateParts = []

  if (!checkFormatting()) {
    return
  }

  var sheet = SpreadsheetApp.getActiveSheet()
  var data = sheet.getDataRange().getValues()

  // loop through each row (excluding 1st row)
  for (var row = 1; row < data.length; row++) {
    
    // add in each row element
    var document = {}
    
    document.partName = data[row][0]
    document.barcode = data[row][1]
    document.datasheet = data[row][2]
    document.isConsumable = data[row][3]
    document.description = data[row][4]
    document.stock = parseInt(data[row][5])
    document.category = data[row][6]
    document.imageUrl = data[row][7]
    document.type = data[row][3] ? "Consumable" : "Returnable"

    // if part with matching barcode found
    if (findPart(document.barcode, apikey)) {
      
      // push the row # + 1 to the array
      duplicateParts.push(row + 1)
      
      // skip current iteration and go to next loop iteration
      continue
    }

    const payload = {
      document: document, 
      collection: collectionName,
      database: databaseName, 
      dataSource: clusterName
    }

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      headers: { "api-key": apikey }
    };

    const response = UrlFetchApp.fetch(insertEndpoint, options);
  }

  if (duplicateParts.length === 0) {
    SpreadsheetApp.getUi().alert("Success! All parts added!")
  }
  else {
    SpreadsheetApp.getUi().alert("Duplicate parts in rows: " + duplicateParts + "\nNon-duplicate items were inserted successfully")
  }
        
}

function clearSheet() {
  var sheet = SpreadsheetApp.getActiveSheet()
  var data = sheet.getDataRange().getValues()
  sheet.getRange(2, 1, data.length, 8).clearContent()
}

function viewParts() {
  const apikey = getAPIKey()

  var sheet = SpreadsheetApp.getActiveSheet()
  var data = sheet.getDataRange().getValues()

  // delete previous data
  clearSheet()

  // create endpoint to find all parts
  const findEndpoint = endpoint + "/action/find"

  const payload = {
    collection: collectionName,
    database: databaseName,
    dataSource: clusterName
  }

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: { "api-key": apikey }
 }

  const response = UrlFetchApp.fetch(findEndpoint, options);

  // parse response object to a JS Object
  const parsedResponse = JSON.parse(response.getContentText())

  const allParts = parsedResponse.documents

  for (var i = 0; i < allParts.length; i++) {

    const part = allParts[i]
    const rowValues = [part.partName, part.barcode, part.datasheet, part.isConsumable, part.description, part.stock, part.category, part.imageUrl]

    sheet.getRange(i+2, 1, 1, rowValues.length).setValues([rowValues])
    
  }

}

function updateParts() {
  const apikey = getAPIKey()
  // add insert action to endpoint
  const updateEndpoint = endpoint + "/action/updateOne"

  var invalidParts = []

  if (!checkFormatting()) {
    return
  }

  var sheet = SpreadsheetApp.getActiveSheet()
  var data = sheet.getDataRange().getValues()

  // loop through each row (excluding 1st row)
  for (var row = 1; row < data.length; row++) {
    
    // add in each row element
    var update = {}
    
    const barcode = data[row][1]

    update.partName = data[row][0]
    update.datasheet = data[row][2]
    update.isConsumable = data[row][3]
    update.description = data[row][4]
    update.stock = parseInt(data[row][5])
    update.category = data[row][6]
    update.imageUrl = data[row][7]
    update.type = data[row][3] ? "Consumable" : "Returnable"

    // if no parts with matching barcode found
    if (!findPart(barcode, apikey)) {
      
      // push the row # + 1 to the array
      invalidParts.push(row + 1)
      
      // skip current iteration and go to next loop iteration
      continue
    }

    const payload = {
      collection: collectionName,
      database: databaseName, 
      dataSource: clusterName,
      filter: {barcode: barcode},
      update: {$set: update},
      upsert: false
    }

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      headers: { "api-key": apikey }
    };

    const response = UrlFetchApp.fetch(updateEndpoint, options);
  }

  if (invalidParts.length === 0) {
    SpreadsheetApp.getUi().alert("Success! All parts updated!")
  }
  else {
    SpreadsheetApp.getUi().alert("Invalid part barcodes in rows: " + invalidParts + "\nOther parts were inserted successfully")
  }
}

function deleteParts() {
  const apikey = getAPIKey()
  var sheet = SpreadsheetApp.getActiveSheet()
  var data = sheet.getDataRange().getValues()

  var invalidBarcodeRows = []

  const deleteEndpoint = endpoint + "/action/deleteOne"

  for (var row = 1; row < data.length; row++) {
    const barcodeToDelete = data[row][0]

    if (!findPart(barcodeToDelete, apikey)) {
      invalidBarcodeRows.push(row+1)
      continue
    }

    // query to search via barcode
    const query = {barcode: barcodeToDelete}

    const payload = {
      filter: query,
      collection: collectionName,
      database: databaseName,
      dataSource: clusterName
    }

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      headers: { "api-key": apikey }
  }

    const response = UrlFetchApp.fetch(deleteEndpoint, options);

  }

  if (invalidBarcodeRows.length === 0) {
    SpreadsheetApp.getUi().alert("All parts with barcodes specified were deleted successfully")
  }
  else {
    SpreadsheetApp.getUi().alert("Invalid barcode in rows: " + invalidBarcodeRows + "\nAll other parts were deleted successfully")
  }
}

