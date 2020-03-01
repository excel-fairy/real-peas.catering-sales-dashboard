var SPREADSHEET = {
    spreadSheet: SpreadsheetApp.getActiveSpreadsheet(),
    sheets: {
        dashboard:{
            sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard"),
            exportFileNameCell: 'A3',
            dateCell: 'C7',
            exportRange: {
                r1: 0,
                r2: 58,
                c1: ColumnNames.letterToColumnStart0('A'),
                c2: ColumnNames.letterToColumnStart0('AI')
            }
        },
        dataValid: {
            sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data valid"),
            exportFolderIdCell: 'B2'
        },
        emailAutomation: {
            sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email automation"),
            recipientCell: 'B2',
            carboCopyCell: 'B3',
            subjectCell: 'B7',
            messageCell: 'B8'
        }
    }
};
