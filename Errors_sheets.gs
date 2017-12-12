///////////////////////////////////////////////////////////////////////////////////////////
//https://docs.google.com/spreadsheets/d/1xWRhAvzJqxWL-PL3eCNBCFdcyYMRhV0KlTUiGgULMsg/edit
//Guarda los esrrores inesperado en una hoja de cálculo especifica.
///////////////////////////////////////////////////////////////////////////////////////////
function saveError(e) {
    try {
        var doc = SpreadsheetApp.openById("1xWRhAvzJqxWL-PL3eCNBCFdcyYMRhV0KlTUiGgULMsg"),
            sheet = doc.getSheetByName("Errors"),

            email = Session.getActiveUser().getEmail(),
            date = new Date(),
            type = e.name,
            msg = e.message,
            file = e.fileName,
            line = e.lineNumber;

        sheet.getRange(sheet.getLastRow()+1, getColIndexByName("Email", sheet)).setValue(email);
        sheet.getRange(sheet.getLastRow(), getColIndexByName("Fecha", sheet)).setValue(date);
        sheet.getRange(sheet.getLastRow(), getColIndexByName("Tipo", sheet)).setValue(type);
        sheet.getRange(sheet.getLastRow(), getColIndexByName("Mensaje", sheet)).setValue(msg);
        sheet.getRange(sheet.getLastRow(), getColIndexByName("Archivo", sheet)).setValue(file);
        sheet.getRange(sheet.getLastRow(), getColIndexByName("Linea", sheet)).setValue(line);
        return;

    } catch (e) {
        return;
    }
}
