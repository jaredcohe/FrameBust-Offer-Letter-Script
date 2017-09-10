// FrameBust Offer Letter Script
// Read more here: https://medium.com/p/b2497b2ef7c7

// ID for template, source, target folder
var templateFileId = "1KLxr8sDwStTyYUXftNz7k6y8YCfizLSpNH0UBFZb944";
var sourceDataFileId = "1HKug8-s_ZjIjKGUvaVLQxU6yrSy4APO3utj4wCW89u8";
var targetFolderId = "0B1F-5fwPmkDoZzRzWkZ6ZWd2bzg";

// Get data as array of arrays
function getDataAsArrayOfArrays() {
    var file = SpreadsheetApp.openById(sourceDataFileId);
    var sheet = file.getSheets()[0];
    var range = sheet.getDataRange();
    var values = range.getValues();
    return values;
}

// Duplicate google doc
function createDuplicateDocument(templateFileId, name) {
    var source = DriveApp.getFileById(templateFileId);
    var newFile = source.makeCopy(name);
    var targetFolder = DriveApp.getFolderById(targetFolderId);
    targetFolder.addFile(newFile)
    return DocumentApp.openById(newFile.getId());
}

// Search a paragraph in the document and replaces it with the generated text 
function replaceText(targetDocumentId, keyword, newText) {
    var targetDocument = DocumentApp.openById(targetDocumentId);
    var targetBody = targetDocument.getBody();
    targetBody.replaceText(keyword, newText);
}

// Main function to run
function main() {
    var allData = getDataAsArrayOfArrays();
    var headers = allData[0];
    var allDataForDocument = allData[allData.length-1];

    // Create the target file, with whatever name you want
    var newTargetFileName = "FrameBust - Offer Letter - " + allDataForDocument[2] + " " + allDataForDocument[3];
    var newTargetFile = createDuplicateDocument(templateFileId, newTargetFileName);
    var newTargetFileId = newTargetFile.getId();

    for(var i=0; i<headers.length; i++) {
        var dataForDocument = allDataForDocument[i].toString();

        // For any numbers you may have to use something like this to ensure a comma goes in
        if ((headers[i] == "#SALARY WITH COMMA BUT WITHOUT DOLLAR SIGN#") || (headers[i] == "#OPTIONS#")) {
            dataForDocument = dataForDocument.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
        }
        replaceText(newTargetFileId, headers[i], dataForDocument);
    }
}