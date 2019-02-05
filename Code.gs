// Important notes:
//
// 1. This code assumes the Flight Sheet is in only one parent folder. If it is in multiple, there will be strange behavior
// 2. There are a bunch of global "constants" at the head of the file that can be changed. Note that some refer to the names of sheets
//    inside existing spreadsheets, so those would have to be changed to match
// 3. This assumes that the Flight Sheet is in a directory with subdirectories for the Archives, Billing Export and Daily PDF's. If these
//    are not sub directories, the code can be easily modified to use the ID's of the directories that should be used for each of these.
//

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Check flight sheet and create totals ...', functionName: 'updateFlightSheet_'},
    {name: 'Save flight sheet invoices and reset ...', functionName: 'saveAndReset_'}
  ];
  spreadsheet.addMenu('Flight Sheet', menuItems);
}

// external file and directory names
var archiveSubDirName = 'Archives';
var exportSubDirName = 'Billing Export';
var pdfSubDirName = 'Daily PDFs';

var weeklyExportSpreadsheetName = 'Weekly Invoice Export';
var targetExportSheetName = "Invoices";

// Fixed references into this Flight Sheet
var spreadsheet = SpreadsheetApp.getActive();
var flightSheet = spreadsheet.getSheetByName('Flight Sheet');
var weeklyExportSheetName = 'Xero Invoices - weekly';
var memberSheetName = 'Club Members';
var tugSheetName = 'Tugs';
var gliderSheetName = 'Club Gliders';

// get the first column which includes location labels
var labelColumn = flightSheet.getRange(1, 1, flightSheet.getLastRow(), 1).getValues();

// Find the critical rows from the flight sheet using labels in column 1
var dateRow = 0;
var tugRow = 0;
var headerRow = 0;
var lastFlightRow = 0;
var insertTimesRow = 0;
var ocHeaderRow = 0;    // oc = other charges
var ocEndRow = 0;
for (i=0; i < labelColumn.length; i++) {
  switch (labelColumn[i][0]) {
    case 'Date': dateRow = i+1;
      break;
    case 'Tug': tugRow = i+1;
      break;
    case 'Flight Sheet': headerRow = i+1; 
      break;
    case 'End Flight Sheet': lastFlightRow = i+1; 
      break;
    case 'Glider Times': insertTimesRow = i+2;
      break;
    case 'Other Charges': ocHeaderRow = i+1; 
      break;
    case 'End Charges': ocEndRow = i+1; 
      break;
  }
}
var firstFlightRow = headerRow+1;
var ocFirstRow = ocHeaderRow + 1;

// Grab the date from the flightSheet
var dateCol = 2;
var flyingDateField = flightSheet.getRange(dateRow, dateCol, 1, 1).getValues()[0,0];
var flyingDate = new Date(flyingDateField);
var resetBackgroundColour = flightSheet.getRange(dateRow, dateCol).getBackground();

// Get the tug
var tugCol = 2;
var tug = flightSheet.getRange(tugRow, tugCol, 1, 1).getValues();

var flightSheetColumns = flightSheet.getLastColumn();
var insertTimesCol = 2;

// Work out the data columns from the flight sheet based on the column header names
var headers = flightSheet.getRange(headerRow, 1, 1, flightSheetColumns).getValues()[0];
var P1NameCol = headers.indexOf('P1');
var P2NameCol = headers.indexOf('P2');
var billToMemberNameCol = headers.indexOf('Bill To Member');
var P1TowCostCol = headers.indexOf('P1 Tow Cost');
var P1GliderCostCol = headers.indexOf('P1 Glider Cost');
var P2TowCostCol = headers.indexOf('P2 Tow Cost');
var P2GliderCostCol = headers.indexOf('P2 Glider Cost');
var billToMemberTowCostCol = headers.indexOf('Bill To Member Tow Cost');
var billToMemberGliderCostCol = headers.indexOf('Bill To Member Glider Cost');
var gliderNameCol = headers.indexOf('Glider');
var flightLaunchCol = headers.indexOf('Launch Time');
var flightLandingCol = headers.indexOf('Landing Time');
var billingCodeCol = headers.indexOf('Billing Code');
var durationCol = headers.indexOf('Duration');
var heightCol = headers.indexOf('Height (100\'s ft)') 
var gfaNumCol = headers.indexOf('Passenger GFA #');

// work out the data columns from the other charges header
headers = flightSheet.getRange(ocHeaderRow, 1, 1, flightSheetColumns).getValues()[0];
var ocTypeCol = headers.indexOf('Type');
var ocMemberCol = headers.indexOf('Member');
var ocAmountCol = headers.indexOf('Amount');
var ocDescriptionCol = headers.indexOf('Description');

//
// updateFlightSheet_()
//

function updateFlightSheet_() {

  // get the rows of the flightSheet
  var flights = flightSheet.getSheetValues(firstFlightRow, 1, lastFlightRow-firstFlightRow+1, flightSheetColumns);
  
  // get the rows of the other charges
  var otherCharges = flightSheet.getSheetValues(ocFirstRow, 1, ocEndRow-ocFirstRow+1, flightSheetColumns);

  // build a dictionary of all the members info
  var membersInfo = {};
  var memberSheet = spreadsheet.getSheetByName(memberSheetName);
  var memberData = memberSheet.getRange(2, 1, memberSheet.getLastRow(), 5).getValues();
  for (var i = 0; i < memberData.length; i++) {
    var name = memberData[i][0];
    if (name != "") {
      membersInfo[name] = {};
      membersInfo[name].invoiceId = memberData[i][1];
      membersInfo[name].homePhone = memberData[i][2];
      membersInfo[name].mobilePhone = memberData[i][3];
      membersInfo[name].email = memberData[i][4];
    }
  }

  // build a dictionary of tug accounting info
  var tugInfo = readAccountInfo_(tugSheetName);
  var gliderInfo = readAccountInfo_(gliderSheetName);
  
  if (checkDataValidity_(flights, otherCharges, membersInfo)) {
    updateGliderTimes_(flights);
    createXeroExports_(flights, otherCharges, membersInfo, tugInfo, gliderInfo);
  }

}

// readAccountInfo_() assumes the key names are in column 1 of the given sheet, and the Xero Inventory Item Code and the Xero Account Code
// are in columns with headers of those names
function readAccountInfo_(sheetName) {
  var invCodeName = "Xero Inventory Item Code";
  var acctCodeName = "Xero Account Code";
  var accountInfo = {};
  var sheet = spreadsheet.getSheetByName(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var invCodeCol = headers.indexOf(invCodeName)+1;
  var acctCodeCol = headers.indexOf(acctCodeName)+1;
  for (row=2; row<=sheet.getLastRow(); row++) {
    var name = sheet.getRange(row, 1).getValue();
    accountInfo[name] = {};
    accountInfo[name].invCode = sheet.getRange(row, invCodeCol).getValue();
    accountInfo[name].acctCode = sheet.getRange(row, acctCodeCol).getValue();
  }
  return accountInfo;
}

function updateGliderTimes_(flights) {  

  // Create dictionary of names with total times
  var gliderTimes = {};
  
  for (var i = 0; i < flights.length; i++) {
    var glider = flights[i][gliderNameCol];
    var flightDuration = (flights[i][flightLandingCol] - flights[i][flightLaunchCol]);    // in milliseconds
    if (glider != "" && glider != "-") {
      if (gliderTimes[glider] == null) {
        gliderTimes[glider] = { time:0, flights:0 };
      };
      gliderTimes[glider].time += flightDuration;
      gliderTimes[glider].flights += 1;
    };
  };

  // create a table to insert
  var gliderTable = [];
  for (var glider in gliderTimes) {
    gliderTable.push([glider, gliderTimes[glider].time/1000/60/60/24, gliderTimes[glider].flights]); // convert to fraction of day for google sheets
  };

  // insert the table
  flightSheet.getRange(insertTimesRow, insertTimesCol, gliderTable.length, 3).setValues(gliderTable);
  flightSheet.getRange(insertTimesRow, insertTimesCol+1, gliderTable.length, 1).setNumberFormat('[hh]:mm');
  
}

function checkDataValidity_(flights, otherCharges, membersInfo) {
  dataOK = true;
  
  // Is there a flying date
  if (flyingDateField == '') {
    showAlert_('Missing Flying Date', 'Please fix and re-run');
    dataOK = false;
  } else {
    // Is the date OK - i.e. before Now, but within the last 7 days?
    now = new Date();
    daysDiff = (now.valueOf() - flyingDate.valueOf()) / (24*60*60*1000);
    if ( daysDiff < 0 || daysDiff > 7 ) {
      showAlert_('Double check the Flying Date', 'It is more than 7 days ago or in the future');
      // allow to continue, as might be fixing old Flight Sheet
    }
  }
  
  // Have we selected a tug?
  if (tug == '') {
    showAlert_('Missing Tug Name', 'Please fix and re-run');
    dataOK = false;
  }
  
  // check each row in the flight data
  for (var i = 0; i < flights.length; i++) {
    
    var glider = flights[i][gliderNameCol];
    var P1Name = flights[i][P1NameCol];
    var P1TowCost = flights[i][P1TowCostCol];    
    var P1GliderCost = flights[i][P1GliderCostCol];     
    var P2Name = flights[i][P2NameCol];
    var P2TowCost = flights[i][P2TowCostCol];
    var P2GliderCost = flights[i][P2GliderCostCol];
    var billToMemberName = flights[i][billToMemberNameCol];
    var billToMemberTowCost = flights[i][billToMemberTowCostCol];
    var billToMemberGliderCost = flights[i][billToMemberGliderCostCol];
    var billingCode = flights[i][billingCodeCol];
    var duration = flights[i][durationCol];
    var height = flights[i][heightCol];
    var gfaNum = flights[i][gfaNumCol];
    
    flightNum = i + 1;
        
    // check each row with a P1 or glider
    if (P1Name != '' || glider != '') {
      
      if (glider == '') {
        showAlert_('Flight ' + flightNum + ': has no glider.', 'Please fix and re-run');
        dataOK = false;
      }
 
      if (P1Name == '') {
        showAlert_('Flight ' + flightNum + ': has no P1.', 'Please fix and re-run');
        dataOK = false;
      }

      if (billingCode == '' || billingCode == '-') {
        showAlert_('Flight ' + flightNum + ': has no billing code.', 'Please fix and re-run');
        dataOK = false;
      }
    
      if (duration == '' ) {
        showAlert_('Flight ' + flightNum + ': has missing duration.', 'Please fix and re-run');
        dataOK = false;
      }
      
      if (height == '' ) {
        showAlert_('Flight ' + flightNum + ': has invalid height.', 'Please fix and re-run');
        dataOK = false;
      }
      
      if (P1TowCost > 0 || P1GliderCost > 0) {
        // check to see if we know about this person for billing. If not alert.
        if (membersInfo[P1Name] == null) {
          showAlert_('Flight ' + flightNum + ': Member "' + P1Name + '" not found in "' + memberSheetName + '" sheet tab, but has flight charges.', 'Please fix and re-run.');
          dataOK = false;
        }
      }
      
      if (P2TowCost > 0 || P2GliderCost > 0) {
        // check to see if we know about this person for billing. If not alert.
        if (membersInfo[P2Name] == null) {
          showAlert_('Flight ' + flightNum + ': Member "' + P2Name + '" not found in "' + memberSheetName + '" sheet tab, but has flight charges.', 'Please fix and re-run.');
          dataOK = false;
        }
      }
      
      if (billToMemberTowCost > 0 || billToMemberGliderCost) {
        // check to see if we know about this person for billing. If not alert.
        if (membersInfo[billToMemberName] == null) {
          showAlert_('Flight ' + flightNum + ': Member "' + billToMemberName + '" not found in "' + memberSheetName + '" sheet tab, but has flight charges.', 'Please fix and re-run.');
          dataOK = false;
        }        
      }

      switch (billingCode) {
        case 'MEM-AEF':
          if (gfaNum == '') {
            showAlert_('Flight ' + flightNum + ': ' + billingCode + ' flight must have Passenger GFA Number.', 'Please fix and re-run');
            dataOK = false;
          }
          if (billToMemberName == '') {
            showAlert_('Flight ' + flightNum + ': ' + billingCode + ' flight must have Bill To Member name.', 'Please fix and re-run');
            dataOK = false;
          }          
          break;
        case 'AEF':
          if (gfaNum == '') {
            showAlert_('Flight ' + flightNum + ': ' + billingCode + ' flight must have Passenger GFA Number.', 'Please fix and re-run');
            dataOK = false;
          }
          break;
        case 'MEM-P2':
        case 'INST-P2':
        case 'MEM-SHR':
        case 'RMC-P2':
        case 'RMO-P2':
          if (P2Name == '') {
            showAlert_('Flight ' + flightNum + ': ' + billingCode + ' flight must have P2 member name.', 'Please fix and re-run');
            dataOK = false;
          }
          break;
        case 'MEM-P1':
        case 'MEM-SHR':
        case 'TO/DP':
        case 'TO/PO':
        case 'RMC-P1':
        case 'RMO-P1':
          if (P1Name == '') {
            showAlert_('Flight ' + flightNum + ': ' + billingCode + ' flight must have P1 member name.', 'Please fix and re-run');
            dataOK = false;
          }
          break;
      }    
    }
  }
  
  return dataOK;
}


function createXeroExports_(flights, otherCharges, membersInfo, tugInfo, gliderInfo) {
  
  // get the two export sheets, creating if necessary
  var weeklyExportSheet = getClearedSheet_(weeklyExportSheetName);
  
  // write the headers
  createXeroExportHeaders_(weeklyExportSheet, membersInfo);

  // write the invoice line items
  createXeroExportRows_(flights, otherCharges, weeklyExportSheet, membersInfo, tugInfo, gliderInfo);
  
}

function getClearedSheet_(sheetName) {
  // Create new sheets if needed
  var exportSheet = spreadsheet.getSheetByName(sheetName);
  if (exportSheet) {
    exportSheet.clear();
    exportSheet.activate();
  } else {
    exportSheet =
      spreadsheet.insertSheet(sheetName, spreadsheet.getNumSheets());
  }
  flightSheet.activate();
  return exportSheet;
}

function createXeroExportHeaders_(exportSheet) {
  // create the headers required for Xero
  var headers = [
    [
     '*ContactName', 
     'EmailAddress', 
     'POAddressLine1', 
     'POAddressLine2', 
     'POAddressLine3', 
     'POAddressLine4', 
     'POCity', 
     'PORegion', 
     'POPostalCode', 
     'POCountry', 
     '*InvoiceNumber', 
     'Reference', 
     '*InvoiceDate', 
     '*DueDate', 
     'InventoryItemCode', 
     '*Description', 
     '*Quantity', 
     '*UnitAmount', 
     'Discount', 
     '*AccountCode', 
     '*TaxType', 
     'TrackingName1', 
     'TrackingOption1', 
     'TrackingName2', 
     'TrackingOption2', 
     'Currency', 
     'BrandingTheme'
    ]
  ];
  exportSheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
}

function createXeroExportRows_(flights, otherCharges, wkExportSheet, membersInfo, tugInfo, gliderInfo) {
  
  // calculate/format the dates - Note this currently enforces australian date format
  var tz = Session.getTimeZone();
  var invDate = flyingDate;
  var dueDate = new Date(invDate.valueOf() + 6*24*60*60*1000);
  var invoiceDate = Utilities.formatDate(invDate, tz, "dd/MM/YYYY");
  var invoiceDueDate = Utilities.formatDate(dueDate, tz, "dd/MM/YYYY");

  // adjust the week # date used for the invoice number so that all days on the weekend are invoiced together using
  // the week# of the Saturday (Bring Monday-1 through Friday-5 fwd and Sunday back)
  var today = invDate.getDay();
  var invDateAdjustment = 0;
  if (today > 0) {
    invDateAdjustment = (6-today);
  } else {
    invDateAdjustment = -1;
  }
  var invBaseDate = new Date(invDate.valueOf() + invDateAdjustment*24*60*60*1000);
  var invoiceWeekBase = Utilities.formatDate(invBaseDate, tz, "YYYY-MM-'WK'W-");
 
  // Buid a dictionary of members with flight costs from today
  memberCost = {};
  for (var i = 0; i < flights.length; i++) {
    addFlightLineItems_(memberCost, invoiceDate, flights[i], tug, tugInfo, gliderInfo, P1NameCol, P1TowCostCol, P1GliderCostCol);
    addFlightLineItems_(memberCost, invoiceDate, flights[i], tug, tugInfo, gliderInfo, P2NameCol, P2TowCostCol, P2GliderCostCol);
    addFlightLineItems_(memberCost, invoiceDate, flights[i], tug, tugInfo, gliderInfo, billToMemberNameCol, billToMemberTowCostCol, billToMemberGliderCostCol);    
  }
  
  // add in other charges
  var name, cost;
  var invCode = 'other';
  var acctCode = '6530';
  for (var i=0; i < otherCharges.length; i++) {
    name = otherCharges[i][ocMemberCol]
    cost = otherCharges[i][ocAmountCol];
    description = otherCharges[i][ocDescriptionCol];
    if (cost > 0) {
      addInvLineItem_(memberCost, name, cost, description, invCode, acctCode);
    }
  }
  
  // write costs to export spreadsheet
  var i = 1;
  for (var name in memberCost) {
    
    var exportSheet = wkExportSheet;
    var invoiceNumber = invoiceWeekBase + membersInfo[name].invoiceId;
    
    for (i=0; i<memberCost[name].lineItems.length; i++) {
      var invoiceItem = memberCost[name].lineItems[i];
      exportSheet.appendRow([
        name, // '*ContactName', 
        membersInfo[name].email, // 'EmailAddress', 
        "", // 'POAddressLine1', 
        "", // 'POAddressLine2', 
        "", // 'POAddressLine3', 
        "", // 'POAddressLine4', 
        "", // 'POCity', 
        "", // 'PORegion', 
        "", // 'POPostalCode', 
        "", // 'POCountry', 
        invoiceNumber, // '*InvoiceNumber', 
        membersInfo[name].invoiceId, // 'Reference', 
        invoiceDate, // '*InvoiceDate', 
        invoiceDueDate, // '*DueDate', 
        invoiceItem.invCode, // 'InventoryItemCode', 
        invoiceItem.description, // '*Description', 
        1, // '*Quantity', 
        invoiceItem.cost, // '*UnitAmount', 
        "", // 'Discount', 
        invoiceItem.acctCode, // '*AccountCode', 
        "GST on Income", // '*TaxType', 
        "", // 'TrackingName1', 
        "", // 'TrackingOption1', 
        "", // 'TrackingName2', 
        "", // 'TrackingOption2', 
        "", // 'Currency', 
        "Flying Charges"  // 'BrandingTheme' 
      ]);
    };
  };
 
  showAlert_('Invoice exports created', '');

}

function addFlightLineItems_(memberCost, invoiceDate, flight, tug, tugInfo, gliderInfo, nameCol, towCostCol, gliderCostCol) {
  var description;
  var glider = flight[gliderNameCol];
  var name = flight[nameCol];
  var towCost = flight[towCostCol]; 
  var gliderCost = flight[gliderCostCol];
  
  if (towCost > 0) {
    description = flightInvLineDescription_(invoiceDate, flight, 'tow');
    addInvLineItem_(memberCost, name, towCost, description, tugInfo[tug].invCode, tugInfo[tug].acctCode);
  };
  
  if (gliderCost > 0) {
    description = flightInvLineDescription_(invoiceDate, flight, 'glider');
    addInvLineItem_(memberCost, name, gliderCost, description, gliderInfo[glider].invCode, gliderInfo[glider].acctCode);
  };
}

function flightInvLineDescription_(invoiceDate, flight, type) {
  var timezone = spreadsheet.getSpreadsheetTimeZone();
  var timeFormat = 'HH:mm';
  var description = invoiceDate + 
    ' ' + Utilities.formatDate(flight[flightLaunchCol], timezone, timeFormat) +
    ' ' + flight[billingCodeCol] + 
    ' ' + type + ' charges for' +
    ' ' + flight[gliderNameCol] + 
    ' Time: ' + Utilities.formatDate(flight[durationCol], timezone, timeFormat) +
    ' Tow Ht: ' + flight[heightCol] + "00";
  return description;
}

function addInvLineItem_(memberCost, name, cost, description, invCode, acctCode) {
  if (memberCost[name] == null ) {
    memberCost[name] = {};
    memberCost[name].lineItems = [];
  };
  var lineItem = {};
  lineItem.cost = cost;
  lineItem.description = description;
  lineItem.invCode = invCode;
  lineItem.acctCode = acctCode;
  memberCost[name].lineItems.push(lineItem);
}

//
// saveAndReset
//
function saveAndReset_() {
  
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert('Is the flight sheet complete and are invoices ready for export?', ui.ButtonSet.YES_NO);
  if (result == ui.Button.NO) {
    return;
  };

  // get Ids for the subdirs based on names. If fixed directories are needed instead, the following lines can be changed to something like:
  // var archiveSubDirId = '<google directory id>';
  var workingDirectoryId = getSpreadSheetParentId_(spreadsheet);  // contains this sheet and the relevant sub-directories
  var archiveSubDirId = getSubDirIdByName_(workingDirectoryId, archiveSubDirName);
  var pdfSubDirId = getSubDirIdByName_(workingDirectoryId, pdfSubDirName);
  var exportSubDirId = getSubDirIdByName_(workingDirectoryId, exportSubDirName);
  
  // references to the export and archive spreadsheets
  var weeklyExportSpreadsheetId = getFileIdByName_(exportSubDirId, weeklyExportSpreadsheetName);
 
  var tz = Session.getTimeZone();
  var saveDate = Utilities.formatDate(flyingDate, tz, "YYYY-MM-dd");

  var weeklySheet = spreadsheet.getSheetByName(weeklyExportSheetName);
    
  // copy the current spreadsheet to one with the date appended and move it to the archive sub dir
  backupFlightSheet_(spreadsheet, saveDate, archiveSubDirId);
  
  // copy the weekly and monthly export invoice line items to the separate weekly and monthly google sheets
  exportXeroInvoices_(weeklySheet, weeklyExportSpreadsheetId);
  
  // create a PDF version of the flight sheet
  saveAsPDF_(spreadsheet, saveDate, pdfSubDirId);
  
  // traverse the flight sheet resetting the contents of all cells with the same background colour as the date cell
  var backgrounds = flightSheet.getRange(1,1,flightSheet.getLastRow(), flightSheet.getLastColumn()).getBackgrounds();
  for (i = 0; i < backgrounds.length; i++) {
    for (j = 0; j < backgrounds[i].length; j++) {
      if (backgrounds[i][j] == resetBackgroundColour) {
        flightSheet.getRange(i+1,j+1).clearContent();
      }
    }
  }

  // delete the glider times rows
  flightSheet.deleteRows(insertTimesRow, flightSheet.getLastRow()-insertTimesRow+1);
  
  // delete the Xero Export Tabs
  spreadsheet.deleteSheet(weeklySheet);
  
  showAlert_('All done!', '');
  
}

function exportXeroInvoices_ (sourceSheet, exportSpreadsheetID) {
  var fullRange = sourceSheet.getDataRange();
  var sourceDataRange = sourceSheet.getRange(2, 1, fullRange.getNumRows()-1, fullRange.getNumColumns());
  var tss = SpreadsheetApp.openById(exportSpreadsheetID); // tss = target spreadsheet
  var ts = tss.getSheetByName(targetExportSheetName); // ts = target sheet
  var targetRange = ts.getRange(ts.getLastRow()+1, 1, sourceDataRange.getNumRows(), sourceDataRange.getNumColumns() );
  targetRange.setValues(sourceDataRange.getValues());
}

function backupFlightSheet_(ss, saveDate, saveDirId) {
  var copyFile = DriveApp.getFileById(ss.copy(ss.getName() + " " + saveDate).getId());
  var parentFolder = copyFile.getParents().next();  // just created so assume only one parent
  DriveApp.getFolderById(saveDirId).addFile(copyFile); // copy it to the archive folder
  parentFolder.removeFile(copyFile);  // remove it from the parent folder
}

function saveAsPDF_(ss, saveDate, saveDirId) {
    var url = Drive.Files.get(ss.getId())
        .exportLinks['application/pdf'];
    url = url + 
      '&gid=0' +    // just the first sheet
      '&size=A4' + //paper size
      '&portrait=false' + //orientation, false for landscape
      '&fitw=true' + //fit to width, false for actual size
      '&sheetnames=false' + 
      '&printtitle=true' + 
      '&pagenumbers=false' + //hide optional
      '&gridlines=true' + //false = hide gridlines
      '&fzr=false'; //do not repeat row headers (frozen rows) on each page
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
        headers: {
            'Authorization': 'Bearer ' + token
        }
    });
    DriveApp.getFolderById(saveDirId).createFile(response.getBlob()).setName(ss.getName() + " " + saveDate);
}

function showAlert_(title, message) {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(title, message, ui.ButtonSet.OK);
  return result;
}

function getSpreadSheetParentId_(ss) {
  var file = DriveApp.getFileById(ss.getId());
  var folders = file.getParents();
  var id = folders.next().getId();      // assume I have at least one parent, and just return the first one
  return id;
}

function getSubDirIdByName_(subdirId, folderName) {
  var folder = DriveApp.getFolderById(subdirId);
  var subFolders =  folder.getFoldersByName(folderName);
  var id = null;
  if (subFolders.hasNext()) {
    id = subFolders.next().getId();    // select the first folder with this name
  } else {
    showAlert_("Internal Error", "Unable to locate folder: " + folderName);
  }
  return id;
}

function getFileIdByName_(subdirId, fileName) {
  var folder = DriveApp.getFolderById(subdirId);
  var files =  folder.getFilesByName(fileName);
  var id = null;
  if (files.hasNext()) {
    id = files.next().getId();    // select the first file with this name
  } else {
    showAlert_("Internal Error", "Unable to locate file: " + fileName);
  }
  return id;
}

/* Revision History

*/