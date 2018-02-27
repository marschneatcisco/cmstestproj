/*
 * Create a new document from the template file
 */
function newLabelTemplate(filename) {
  //create template of label file
  var labelTemplateFile = DriveApp.getFileById(DocumentApp.openByUrl("https://docs.google.com/document/d/1rLpp1hhFASftN5VvGx2VFz_fKE2WoNqEhF2cJxW5YhI/edit").getId());
  var labelsFolder = DriveApp.getFoldersByName("ff_labels").next();

  // TODO: verify file doesn't exist before we try to setName?
  var editableLabelDocId = labelTemplateFile.makeCopy(labelsFolder).setName(filename).getId();
  var editableLabelDoc = DocumentApp.openById(editableLabelDocId);
  //TODO: verify file exists before returning?
  return editableLabelDoc;
}

function formatLabelFromSquare(orderDetails, customerName, totalMeals, totalSoups) {
  var body = [];
  var mealCount = 1;
  orderDetails.itemizations.forEach( function(item) {
    // if this is a soup, don't print a separate label
    //TODO: what if there are only soups?
    // if totalMeals == 0 && totalSoups > 0
    if (item.name == "Clam Chowder Soup")
      return;

    for (var c = 1; c <= item.quantity; c++){
      var labelString = customerName + "\t" + "Meal " + mealCount + " of " + totalMeals + "\n";

      if (item.item_variation_name != "")
        labelString += item.name + " (" + item.item_variation_name + ")\n";
      else
        labelString += item.name + "\n"

      labelString += "Side: " + item.modifiers[0].name + "\n";
      if (totalSoups > 0)
        labelString += totalSoups + " soups in order";

      body.push(labelString);
    }
    mealCount++;
  });
  // Last Name       Meal X of Y
  // Hand Breaded Fish (ADULT|CHILD)
  // Side: [Fries|Red Potato]
  // Z Soups in Order

  // if another meal, new "page"
  return body;
}

function formatLabelFromSheet(orderDetails) {
  //TODO: format from data available in Sheet
  return ['test body when not generated from squqre'];
}

function createLabelFile(orderNumber, orderDetails, customerName, totalMeals, totalSoups) {
  var editableLabelDoc = newLabelTemplate("Order " + orderNumber + ": " + customerName);
  //for each meal, enter into label

  var body = editableLabelDoc.getBody();
  var text = formatLabelFromSquare(orderDetails, customerName, totalMeals, totalSoups);
  for (var line in text) {
    body.appendParagraph(text[line]);
    body.appendPageBreak();
  }

  return editableLabelDoc.getUrl();

}
/*
 * Create label from Sheet data
 */
function createLabelFileFromSheet(orderDetails) {
  //As Order Number and Last name should be globally unique, this should make it easy to find in the Drive folder
  var editableLabelDoc = newLabelTemplate("Order " + orderDetails['Order Number'] + ": " + orderDetails['Last Name']);
  var body = editableLabelDoc.getBody();
  var text = formatLabelFromSheet(orderDetails);
  for (var line in text) {
    body.appendParagraph(text[line]);
    body.appendPageBreak();
  }
  return editableLabelDoc.getUrl();
}

function printLabelFromFile(filename_url) {
  // Kathy's label printer print ID.  Probably should get passed to this function
  var printerID = "95bc0f2e-5304-762c-5cca-3508d758c0fe";

  // verify filename exists and we have access to it
  var file = null;
  try {
    file = DriveApp.getFileById(DocumentApp.openByUrl(filename_url).getId());
  } catch (e) {
    Logger.log(e);
    //TODO: why would it ever get to this state, and we should be returning 'false'
    return true;
  }
  var printSuccessful = false;

  var docName = file.getName();
  var docId = file.getId();

  var ticket = {
    version: "1.0",
    print: {
      color: {
        type: "STANDARD_COLOR",
        vendor_id: "Color"
      },
      duplex: {
        type: "NO_DUPLEX"
      }
    }
  };

  var payload = {
    "printerid" : printerID,
    "title"     : docName,
    "content"   : docId,
    "contentType": "google.kix",
    "ticket"    : JSON.stringify(ticket)
  };

  var response = UrlFetchApp.fetch('https://www.google.com/cloudprint/submit', {
    method: "POST",
    payload: payload,
    headers: {
      Authorization: 'Bearer ' + getCloudPrintService().getAccessToken()
    },
    "muteHttpExceptions": true
  });

  response = JSON.parse(response);

  if (response.success) {
    printSuccessful = true;
    Logger.log("%s", response.message);
  } else {
    Logger.log("Error Code: %s %s", response.errorCode, response.message);
  }

  return printSuccessful;
}



/*
 * Functions to get/check OAuth2 stuff as per https://github.com/googlesamples/apps-script-oauth2
 * Example: https://ctrlq.org/code/20061-google-cloud-print-with-apps-script
 */
function showURL() {
  var cpService = getCloudPrintService();
  if (!cpService.hasAccess()) {
    Logger.log(cpService.getAuthorizationUrl());
  }
}

function getCloudPrintService() {
  return OAuth2.createService('print')
    .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
    .setTokenUrl('https://accounts.google.com/o/oauth2/token')
    .setClientId('622309579105-04e1ff61555p28tjg0vcfur0vhr2gotv.apps.googleusercontent.com')
    .setClientSecret('a7I6MOBOgk2v689D8P0xiNEv')
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('https://www.googleapis.com/auth/cloudprint')
    .setParam('login_hint', Session.getActiveUser().getEmail())
    .setParam('access_type', 'offline')
    .setParam('approval_prompt', 'force');
}

function authCallback(request) {
  var isAuthorized = getCloudPrintService().handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('You can now use Google Cloud Print from Apps Script.');
  } else {
    return HtmlService.createHtmlOutput('Cloud Print Error: Access Denied');
  }
}

function getPrinterList() {

  var response = UrlFetchApp.fetch('https://www.google.com/cloudprint/search', {
    headers: {
      Authorization: 'Bearer ' + getCloudPrintService().getAccessToken()
    },
    muteHttpExceptions: true
  }).getContentText();

  var printers = JSON.parse(response).printers;

  for (var p in printers) {
    Logger.log("%s %s %s", printers[p].id, printers[p].name, printers[p].description);
  }
}


