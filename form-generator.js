function onFormSubmit() {
    console.log('Form has been submitted');
    console.log('Script starting...');

    const form = FormApp.getActiveForm();
    const formResponse = form.getResponses().pop()
    const respondentEmail = formResponse.getRespondentEmail();

    if(!AdminLib.getUser(respondentEmail).isDelegatedAdmin) {
        console.error(`Non admin submitted the form: ${respondentEmail}`);

        GmailApp.sendEmail(respondentEmail, 'Not an Admin', 'Your account is not an admin, if this is a mistake, please contact your supervisor');
        form.deleteResponse(formResponse.getId());
        return;
    }

    console.log('Creating variables...');

    // IMPORTANT: Item IDs for form questions can be found when inspecting a question with the Inspector tool
    // you have to look for the 'data-observe-id="131658212"' field
    // the number is the ID

    const userId = formResponse.getResponseForItem(form.getItemById('item_id')).getResponse().toString().trim();
    var user;
    try {
        user = userId ? AdminLib.getUser(userId) : userId;
    } catch(e) {
        console.error(e);

        GmailApp.sendEmail(respondentEmail, "Invalid User ID", `The user id:  ${userId}  is not valid, please try again`);
        form.deleteResponse(formResponse.getId());
        return;
    }

    console.log(`User: ${user.primaryEmail}`);

    const deviceSerial = formResponse.getResponseForItem(form.getItemById('item_id')).getResponse().toString().trim();
    const deviceAsset = formResponse.getResponseForItem(form.getItemById('item_id')).getResponse().toString().trim();
    var device;
    try {
        device = AdminLib.getDevice(deviceSerial || deviceAsset);
    } catch(e) {
        console.error(e);

        GmailApp.sendEmail(respondentEmail, "Invalid Device Serial/Asset", `The script was unable to find the device using: ${device || deviceAsset}`);
        form.deleteResponse(formResponse.getId());
        return;
    }

    console.log(`Device Serial: ${device.serialNumber}`);
    console.log(`Device Asset: ${deviceAsset}`);

    const spreadSheetId = 'spreadsheet_id';
    const infoSpreadSheet = SpreadsheetApp.openById(spreadSheetId);
    var schoolInfo, schoolAbbrev;

    if(user) {
        schoolAbbrev = user?.orgUnitPath.split('/').pop().toUpperCase();

        try {
            schoolInfo = infoSpreadSheet.getSheetByName("School Info").getRange("A2:D").getValues().find(row => row[1] == schoolAbbrev);
            schoolInfo.length;
        } catch(e) {
            console.error(e);

            GmailApp.sendEmail(respondentEmail, "Invalid/Missing School Data", `The script was unable to find the desired school as it is either missing or has invalid data`);
            form.deleteResponse(formResponse.getId());
            return;
        }
    } else {
        schoolAbbrev = 'General Repair'
    }

    console.log(`School abbreviation: ${schoolAbbrev}`)

    const ensureAssigned = !!formResponse.getResponseForItem(form.getItemById('item_id'));
    const deviceParts = form.getItemById('item_id');
    const deivcePartsResponse = formResponse.getResponseForItem(deviceParts)?.getResponse();

    const unformattedDate = new Date();
    const currentSchoolYear = getCurrentSchoolYear(unformattedDate);
    const formattedDate = Utilities.formatDate(unformattedDate, 'UTC+5', 'yyyy-MM-dd');

    const extraNotes = formResponse.getResponseForItem(form.getItemById('item_id')).getResponse();

    console.log(`Today's Date: ${formattedDate}`);
    console.log(`Current School Year: ${currentSchoolYear}`);

    var partsReplaced, partsCost, schema = {};
    var totalCost = 0;

    if(deivcePartsResponse) {
        partsReplaced = deivcePartsResponse.map((val, i) =>
            val?.includes(deviceParts.asCheckboxGridItem().getColumns()[0]) ?
                deviceParts.asCheckboxGridItem().getRows()[i] : null).filter(val => val);

        partsCost = deivcePartsResponse.map((val, i) =>
            val?.includes(deviceParts.asCheckboxGridItem().getColumns()[1]) ?
                infoSpreadSheet.getSheetByName("Device Parts").getRange(i + 2, 2).getValue() : null).filter(val => val);

        totalCost = partsCost?.length ? partsCost.reduce((total, val) => total + val) : 0;
    }

    console.log(`Assign device to user? ${ensureAssigned}`)
    console.log(`Parts replaced on device: ${partsReplaced}`)
    console.log(`Total part cost: ${totalCost}`)

    console.log('Variables have been created');

    if(ensureAssigned) {
        schema.annotatedUser = user ? user.primaryEmail : user;
        schema.orgUnitPath = user ? user.orgUnitPath.replace('users', 'devices') : 'org_path';
        schema.annotatedLocation = user ? schoolAbbrev : '';
    }

    if(deviceAsset) schema.annotatedAssetId = deviceAsset;
    if(partsReplaced) schema.notes = `${user ? `Broken by ${user.primaryEmail.split('@')[0]}` : schoolAbbrev} on ${formattedDate}: ${partsReplaced.join('/')} ${device.notes ? `\n${device.notes}` : ''}`;

    AdminLib.updateDeviceCustomFields(schema, device);

    if(!partsReplaced) {
        console.log('Script has ended successfully!')
        return;
    }

    if(partsCost?.length && user) {
        let invoice = createInvoice(
            device,
            user,
            formattedDate,
            schoolInfo,
            partsReplaced,
            partsCost,
            totalCost
        );

        moveFile(invoice.getId(), "Invoices");
        moveFile(invoice.getId(), currentSchoolYear);
        moveFile(invoice.getId(), schoolAbbrev);

        GmailApp.sendEmail(
            `${respondentEmail}, ${schoolInfo[3]}`, /* Email list */
            "Student Invoice",                      /* Email Title */
            `This is an automated invoice, the student ${user.name.fullName} has broken their device. Please contact ${respondentEmail} with any questions`,      /* Email body */
            {attachments: invoice}                  /* Email attachment */
        );
    }

    console.log("Email sent successfully!");

    var spreadSheet, sheet;
    let spreadSheetName = `${currentSchoolYear} Device Damage Spreadsheet`;

    try {
        console.log(`Attempting to open the spreadsheet named: ${spreadSheetName}`);
        spreadSheet = SpreadsheetApp.openById(DriveApp.getFilesByName(spreadSheetName).next().getId());
    } catch(e) {
        console.error(e);
        console.log('No spreadsheet found');

        console.log(`Creating spreadsheet named: ${spreadSheetName}...`);
        spreadSheet = SpreadsheetApp.create(spreadSheetName);
        console.log('Spreadsheet created successfully');

        DriveApp.getFileById(spreadSheet.getId()).moveTo(DriveApp.getFolderById("drive_folder_id"));

        moveFile(spreadSheet.getId(), "Invoices");
        moveFile(spreadSheet.getId(), currentSchoolYear);
    }

    console.log('Spreadsheet opened successfully');

    try {
        console.log(`Attempting to get sheet name: ${schoolAbbrev}`);
        sheet = spreadSheet.getSheetByName(schoolAbbrev).activate();

    } catch(e) {
        console.error(e);
        console.log('No sheet found');

        console.log(`Creating sheet named: ${schoolAbbrev}`);
        sheet = spreadSheet.insertSheet().setName(schoolAbbrev).activate();
        sheet.appendRow(["Date", "User ID", "Last Name", "First Name", "Device Serial", "Device Asset", "Parts Replaced", "Total Cost", "Notes"])
        sheet.getRange('A1:I1').setBackground('#d2a3d3').setFontSize(11).setTextStyle(SpreadsheetApp.newTextStyle().setBold(true).build());
        sheet.setFrozenRows(1);

        sheet.setColumnWidth(3, 125);
        sheet.setColumnWidth(4, 125);
        sheet.setColumnWidth(5, 115);
        sheet.setColumnWidth(6, 115);
        sheet.setColumnWidth(7, 425);
        sheet.setColumnWidth(9, 425);

        try {
            spreadSheet.deleteSheet(spreadSheet.getSheetByName("Sheet1"));
        } catch(e) {}

        console.log('Sheet created and formatted');
    }

    console.log('Inserting information into spreadsheet...');

    sheet.appendRow([
        formattedDate,
        user ? user.primaryEmail.split('@')[0] : '',
        user ? user.name.familyName : '',
        user ? user.name.givenName : '',
        device.serialNumber,
        deviceAsset ? deviceAsset : device.annotatedAssetId,
        partsReplaced.join(' / '),
        partsCost?.length ? totalCost : '',
        extraNotes
    ]);

    console.log('Script has ended successfully!');
}

function moveFile(fileID, destination) {
    let file = DriveApp.getFileById(fileID);
    try {
        console.log(`Attempting to move file to the folder: ${destination}...`);
        file.moveTo(file.getParents().next().getFoldersByName(destination).next());
    } catch (e) {
        console.error(e);
        console.log('No folder found');

        console.log(`Creating folder: ${destination}...`)
        file.getParents().next().createFolder(destination);

        console.log(`Attempting to move file to the folder: ${destination}...`);
        file.moveTo(file.getParents().next().getFoldersByName(destination).next());
    }

    console.log('File moved successfully');
}

/**
 * @param {Object} device
 * @param {Object} user
 * @param {Date} date
 * @param {Array} schoolInfo
 * @param {Array} deviceDamage
 * @param {Array} priceArray
 * @param {string} totalCost
 * @returns {string} Returns the ID of the Invoice document
 **/

function createInvoice(device, user, date, schoolInfo, deviceDamage, priceArray, totalCost) {
    console.log('Cloning the template invoice...')
    let templateInvoiceId = 'google_doc_id';

    let invoice = DocumentApp.openById(DriveApp.getFileById(templateInvoiceId).makeCopy("Temp Invoice").getId());
    let invoiceBody = invoice.getBody();

    console.log("Replacing text in invoice...")
    invoiceBody.replaceText('{Date}', date);
    invoiceBody.replaceText('{School}', schoolInfo[0]);
    invoiceBody.replaceText('{Address}', schoolInfo[2]);
    invoiceBody.replaceText('{ID}', user.primaryEmail.split('@')[0]);
    invoiceBody.replaceText('{Name}', user.name.fullName);
    invoiceBody.replaceText('{Serial}', device.serialNumber);
    invoiceBody.replaceText('{Model}', device.model);
    invoiceBody.replaceText('{Issues}', deviceDamage.join(' / '));
    invoiceBody.replaceText('{Price}', priceArray.join(' / $'));
    invoiceBody.replaceText('{TotalPrice}', totalCost);
    console.log("Texted replaced successfully")

    console.log("Saving and renaming invoice...")
    invoice.saveAndClose();
    invoice.setName(`${date} - ${user.primaryEmail.split('@')[0]} - Invoice`);
    console.log("Invoice saved and renamed successfully")

    return invoice;
}

function getCurrentSchoolYear(date) {
    const currentYear = date.getFullYear();

    // Determine the school year range
    const schoolYearStart = date.getMonth() < 6 // July is 6 because Jan is 0
        ? new Date(`7/1/${currentYear - 1}`)      // July 1st of last year
        : new Date(`7/1/${currentYear}`);         // July 1st of this year

    const schoolYearEnd = date.getMonth() < 6   // July is 6 because Jan is 0
        ? new Date(`6/30/${currentYear}`)         // June 30th of this year
        : new Date(`6/30/${currentYear + 1}`);    // June 30th of next year

    // Compare the date with the school year range
    // return date >= schoolYearStart && date <= schoolYearEnd;
    return `${schoolYearStart.getFullYear()}-${schoolYearEnd.getFullYear()}`;
}