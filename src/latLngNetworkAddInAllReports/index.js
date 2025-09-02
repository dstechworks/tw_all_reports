const { google } = require('googleapis');
const { Pool } = require('pg');
const path = require("path");
const xlsx = require("xlsx");
const { getCredentialsPath } = require('../utility/pathUtils');

const pool = new Pool({
    user: "postgres",
    host: 'db.mgampbhmlnalxohuobpr.supabase.co',
    database: "postgres",
    password: 'gplVhDuxLDMeBKxs',
    port: 5432,
});

// Load workbook
// const inputFilePath = path.join(__dirname, "reports", "Techworks Assets.xlsx");
const inputFilePath = path.join(__dirname, "reports", "Techworks Assets_updated_3.xlsx");
const workbook = xlsx.readFile(inputFilePath);
// Get sheet
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
// Convert to JSON
let sheetData = xlsx.utils.sheet_to_json(sheet, { defval: "" });

// GOOGLE API VARIABLES
const digiQuadSpreadsheetId = "1gZGiljfrwMCwNlESdhZu2w0Ff5iwSvGEd6t2r3ZC7js";
const backwallSpreadsheetId = "1bu-l2ds0tO51IHxAMw13HPQFQKegLBVWPXV2fTmRI9I";
let workbookData = {};


async function getDataFromGoogleSheets(sheetID, reference) {
    const initializeGoogleSheetsAPI = async () => {
        try {
            // Initialize the authentication client
            const auth = new google.auth.GoogleAuth({
                keyFile: getCredentialsPath(),
                scopes: ["https://www.googleapis.com/auth/spreadsheets"],
            });

            // Get the authenticated client
            const authClientObject = await auth.getClient();

            // Create the Sheets instance
            const sheets = google.sheets({ version: 'v4', auth: authClientObject });

            return sheets; // Return the sheets instance
        } catch (error) {
            console.error("Error initializing Google Sheets API:", error);
            throw error;
        }
    };

    const fetchSheetNames = async (sheets) => {
        try {
            // Get workbook names present in the spreadsheet
            const response = await sheets.spreadsheets.get({
                spreadsheetId: sheetID,
            });

            const sheetNames = response.data.sheets.map(sheet => sheet.properties.title);
            // console.log('\n');
            // console.log('Sheet Names:', sheetNames);

            return sheetNames; // Return the sheet names
        } catch (error) {
            console.error("Error fetching workbook names:", error);
            throw error;
        }
    };

    const fetchSheetData = async (sheets, sheetNames) => {
        try {
            for (let i = 0; i < sheetNames.length; i++) {
                const sheetName = sheetNames[i];
                // Fetch data for each sheet
                const response = await sheets.spreadsheets.values.get({
                    spreadsheetId: sheetID,
                    range: sheetName,
                });

                const data = response.data.values || [];
                // console.log(`Data for ${sheetName}:`, data.length);

                // Change array of array data to array of objects like API response
                const [headers, ...rows] = data;
                let result = rows.map(row => Object.fromEntries(headers.map((key, index) => [key, row[index]])));
                workbookData[sheetName] = result;
            }
        } catch (error) {
            console.error("Error fetching data for sheets:", error);
            throw error;
        }

        return true;
    };

    try {
        let sheets = await initializeGoogleSheetsAPI();
        let sheetNames = await fetchSheetNames(sheets);
        let getWorkbookRes = await fetchSheetData(sheets, sheetNames);
        return getWorkbookRes;
    } catch (error) {
        console.error("Error during Google Sheets data retrieval:", error);
    }
}

function isValidNumber(n) {
    return typeof n === "number" && !isNaN(n);
}

function formatCoordinate(value) {
    if (!isValidNumber(value)) return null;
    return Number(value.toFixed(7)); // round to 7 decimal places
}

async function main() {
    // // Fetch data from both Google Sheets
    // console.log("Fetching data from DigiQuad spreadsheet...");
    // await getDataFromGoogleSheets(digiQuadSpreadsheetId, "digiQuad");

    // console.log("Fetching data from Backwall spreadsheet...");
    // await getDataFromGoogleSheets(backwallSpreadsheetId, "backwall");

    // const dqData = workbookData["DQ"];
    // const backwallData = workbookData["Main Sheet Backwall"];


    // // console.log(dqData);
    // // console.log(backwallData);

    // sheetData.forEach(element => {
    //     const sheetDeviceId = element['Device Id'];
    //     const findByDeviceIdInTwTable = dqData.find(item => item['TECKWORKS'] == sheetDeviceId);

    //     if (findByDeviceIdInTwTable) {
    //         element['City'] = findByDeviceIdInTwTable['CITY'];
    //         element['Sim Card Number'] = findByDeviceIdInTwTable['SIM CARD NUMBER'];
    //         element['Network Operator (SIM Cards)'] = findByDeviceIdInTwTable['SIM CARD PROVIDER'];
    //     }

    //     const findByDeviceIdInBackwallTable = backwallData.find(item => item['TECKWORKS ID'] == sheetDeviceId);
    //     if (findByDeviceIdInBackwallTable) {
    //         element['City'] = findByDeviceIdInBackwallTable['CITY'];
    //         element['Sim Card Number'] = findByDeviceIdInBackwallTable['SIM CARD NUMBER'];
    //         element['Network Operator (SIM Cards)'] = findByDeviceIdInBackwallTable['SIM CARD PROVIDER'];
    //     }

    //     // console.log(element);
    // });

    // // Create a new workbook
    // const newWorkbook = xlsx.utils.book_new();

    // // Convert updated JSON back to worksheet
    // const updatedSheet = xlsx.utils.json_to_sheet(sheetData, { header: Object.keys(sheetData[0]) });

    // // Append sheet into new workbook
    // xlsx.utils.book_append_sheet(newWorkbook, updatedSheet, "Techworks Assets");

    // // Dynamic output path (new file, not overwrite)
    // const outputFilePath = path.join(__dirname, "reports", "Techworks Assets_updated_3.xlsx");

    // // Write new file
    // xlsx.writeFile(newWorkbook, outputFilePath);

    // console.log(`✅ New Excel file created: ${outputFilePath}`);


    const response = await pool.query(`select * from tab_device_records`);
    let twTableData = response.rows;

    // Your update logic
    sheetData.forEach(element => {
        const sheetDeviceId = element['Device Id'];
        const findByDeviceIdInTwTable = twTableData.find(item => item['device_id'] == sheetDeviceId);


        if (findByDeviceIdInTwTable) {
            const lat = parseFloat(findByDeviceIdInTwTable['latitude']);
            const lng = parseFloat(findByDeviceIdInTwTable['longitude']);

            if (isValidNumber(lat) && isValidNumber(lng)) {
                const cleanLat = formatCoordinate(lat);
                const cleanLng = formatCoordinate(lng);
                element['Latitude Longitude Details'] = `${cleanLat},${cleanLng}`;
                element['City'] = findByDeviceIdInTwTable['city'];
                element['Sim Card Number'] = findByDeviceIdInTwTable['sim_details'];
            } else {
                // Put empty string if invalid
                element['Latitude Longitude Details'] = "";
                element['City'] = "";
                element['Sim Card Number'] = "";
            }

            element['Network Operator (SIM Cards)'] = findByDeviceIdInTwTable['sim_card_provider'] || "";
        }
    });

    // Create a new workbook
    const newWorkbook = xlsx.utils.book_new();

    // Convert updated JSON back to worksheet
    const updatedSheet = xlsx.utils.json_to_sheet(sheetData, { header: Object.keys(sheetData[0]) });

    // Append sheet into new workbook
    xlsx.utils.book_append_sheet(newWorkbook, updatedSheet, "Techworks Assets");

    // Dynamic output path (new file, not overwrite)
    const outputFilePath = path.join(__dirname, "reports", "Techworks Assets_updated_new.xlsx");

    // Write new file
    xlsx.writeFile(newWorkbook, outputFilePath);

    console.log(`✅ New Excel file created: ${outputFilePath}`);
}

main();