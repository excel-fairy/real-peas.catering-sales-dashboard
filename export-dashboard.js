function exportDashboard() {
    // Protect the sheet while the export is in progress
    var protection = protectDashboardSheet();
    setDate();

    var exportOptions = {
        sheetId: SPREADSHEET.sheets.dashboard.sheet.getSheetId(),
        exportFolderId: getExportFolderId(),
        exportFileName: getExportFileName(),
        range: SPREADSHEET.sheets.dashboard.exportRange,
        repeatHeader: true,
        portrait: false,
        fileFormat: 'pdf',
        margin: {
            top: 0,
            left: 0,
            right: 0,
            bottom: 0
        }
    };

    var pdfFile = ExportSpreadsheet.apply(exportOptions);
    sendEmail(pdfFile);

    // Unprotect the sheet now the export is done
    protection.remove();
}

function sendEmail(attachment) {
    var emailData = getEmailData();
    var emailAddress = emailData.recipient;
    var subject = emailData.subject;
    var message = emailData.message;
    var carbonCopyEmailAddresses = emailData.carbonCopy;

    var emailOptions = {
        attachments: [attachment.getAs(MimeType.PDF)],
        name: 'Automatic Catering dashboard mail sender',
        cc: carbonCopyEmailAddresses};

    MailApp.sendEmail(emailAddress, subject, message, emailOptions);
}

function getEmailData() {
    return {
        recipient: SPREADSHEET.sheets.emailAutomation.sheet.getRange(
            SPREADSHEET.sheets.emailAutomation.recipientCell).getValue(),
        carbonCopy: SPREADSHEET.sheets.emailAutomation.sheet.getRange(
            SPREADSHEET.sheets.emailAutomation.carboCopyCell).getValue(),
        subject: SPREADSHEET.sheets.emailAutomation.sheet.getRange(
            SPREADSHEET.sheets.emailAutomation.subjectCell).getValue(),
        message: SPREADSHEET.sheets.emailAutomation.sheet.getRange(
            SPREADSHEET.sheets.emailAutomation.messageCell).getValue()
    };
}

function getExportFolderId() {
    return SPREADSHEET.sheets.dataValid.sheet.getRange(SPREADSHEET.sheets.dataValid.exportFolderIdCell).getValue();
}

function getExportFileName() {
    return SPREADSHEET.sheets.dashboard.sheet.getRange(SPREADSHEET.sheets.dashboard.exportFileNameCell).getValue();
}

function setDate() {
    var yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    var formattedDate = Utilities.formatDate(yesterday, 'Australia/Sydney', 'dd/MM/yyyy');
    SPREADSHEET.sheets.dashboard.sheet.getRange(SPREADSHEET.sheets.dashboard.dateCell).setValue(formattedDate);
}

/**
 * Protect the dashboard sheet (and removes everyone from the edit whitelist) while it is being exported
 * (prevent any user to update it)
 */
function protectDashboardSheet() {
    var protection = SPREADSHEET.sheets.dashboard.sheet.protect().setDescription('Export sheet protection');
    protection.removeEditors(protection.getEditors());
    return protection;
}