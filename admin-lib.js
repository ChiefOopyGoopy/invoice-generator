/**
 * THIS LIBRARY HOLDS FUNCTIONS THAT ARE COMMON/REUSED OFTEN IN MY SCRIPTS
 **/

const customerID = 'admin_customer_id';

/**
 * @param {JSON} schema -
 * @param {Object} device -
**/
function updateDeviceCustomFields(schema, device) {
    console.log(`Updating device fields...`);

    console.log(JSON.stringify(schema));

    AdminDirectory.Chromeosdevices.update(
        schema,
        customerID,
        device.deviceId
    );

    console.log('Device fields have been updated')
}

/**
 * @param {string} identifier - device serial/asset ID
 **/
function getDevice(identifier) {
    console.log(`Fetching device from Google Admin using the identifier: ${identifier}`)
    let deviceId = (AdminDirectory.Chromeosdevices.list(customerID, {projection: 'BASIC', query: identifier})).chromeosdevices[0].deviceId;

    return AdminDirectory.Chromeosdevices.get(customerID, deviceId);
}

/**
 * Retrieves a user from Google Admin
 *
 * @param {string} userID - either studentID or teacher email
 * @returns {Object} The user from Google Admin
 **/
function getUser(userID) {
    userID = userID?.toString().split('@')[0]

    console.log(`Fetching user from Google Admin using the identifier: ${userID}`)
    try {
        return AdminDirectory.Users.get(`${userID}@examplestudent.org`)
    } catch {
        return AdminDirectory.Users.get(`${userID}@examplestaff.org`)
    }
}

