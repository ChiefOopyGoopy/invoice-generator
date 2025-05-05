function onOpen() { // Runs when Google Sheet is opened
    const ui = SpreadsheetApp.getUi();

    // Creates a drop-down menu
    ui.createMenu('Upadte Form Data')
        .addItem('Update Data', 'updateForm')
        .addToUi();
}

function updateForm() { // Updates the Form when called
    FormApp.openById('sheet_id').getItemById('item_id').asCheckboxGridItem().setRows(
        // Gets the list of device damage types from the sheet
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Device Damage').getRange('A2:A').getValues().filter(String)
    );

    SpreadsheetApp.getUi().alert('Form data has been updated!');
}