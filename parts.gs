const endpoint = "https://data.mongodb-api.com/app/data-rapuj/endpoint/data/v1"
const clusterName = "Cluster1"
const databaseName = "partsDB"
const collectionName = "parts"

const columnVals = ["Part Name", "Barcode", "Datasheet", "Is Consumable", "Description", "Stock", "Category", "Image URL"]

function getAPIKey() {
  var result = SpreadsheetApp.getUi().prompt(
   'Enter API Key',
   'Key:', SpreadsheetApp.getUi().ButtonSet);
  const apikey = result.getResponseText()
  return apikey;
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

function testFindPart() {
  Logger.log(findPart(1234567890, "8c3rWXDyxEZnX5e1kPRjtfbYMvVw661DfjJynJzKGeJKPX4nA15Mkxw3Oj4sNTHg"))
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
    return false
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

