// FrameBust Offer Letter Script
// Read more here: https://opsmba.com/2017/09/12/how-to-automate-filling-out-google-doc-templates/
// Or here: https://docs.google.com/document/d/15d7_HJ1lPpRqoaF41JH2vTjalu-SlyzGWN2MCcdDrhc/edit?usp=sharing

// ID for template, source, target folder
var templateFileId = "INSERT TEMPLATE FILE ID";
var sourceDataFileId = "INSERT SPREADSHEET FILE ID";
var targetFolderId = "INSERT GOOGLE DRIVE FOLDER ID FOR WHERE TO PUT OUTPUT";

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
