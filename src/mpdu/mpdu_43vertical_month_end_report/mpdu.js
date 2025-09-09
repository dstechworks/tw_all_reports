const { Pool } = require('pg');
const XLSX = require('xlsx');
const { google } = require('googleapis');
const nodemailer = require("nodemailer");
const moment = require('moment-timezone');
const fs = require('fs');
const path = require('path');
const { getCredentialsPath } = require('../../utility/pathUtils');

// Define the folder path for saving the .xlsx files
const reportsFolderPath = path.join(__dirname, 'reports');

const transporter = nodemailer.createTransport({
    host: "smtp.dreamhost.com",
    port: 465,
    secure: true,
    auth: {
        // TODO: replace user and pass values from <https://forwardemail.net>
        user: 'reports@techworks.co.in',
        pass: 'vucto0-socWiz-cifjaj'
    }
});
const pool = new Pool({
    user: "postgres",
    host: 'db.mgampbhmlnalxohuobpr.supabase.co',
    database: "postgres",
    password: 'gplVhDuxLDMeBKxs',
    port: 5432,
});

// GOOGLE API VARIABLES
const spreadsheetId = "17ADQ1OvzA2KhHe1TG5eoCdFDqkHFiumuwIdy9jQ3s2M";
let workbookData = {};

const workbook = XLSX.utils.book_new();
const workbookWBHO = XLSX.utils.book_new();
const workbookWNAG = XLSX.utils.book_new();
const workbookWVIZ = XLSX.utils.book_new();
const workbookWHYD = XLSX.utils.book_new();
const workbookWBLR = XLSX.utils.book_new();
const workbookWCHE = XLSX.utils.book_new();
const workbookWJPR = XLSX.utils.book_new();
const workbookWMUM = XLSX.utils.book_new();
const workbookWPUN = XLSX.utils.book_new();
const workbookWLUC = XLSX.utils.book_new();
const workbookWEUP = XLSX.utils.book_new();
const workbookWCAL = XLSX.utils.book_new();
const workbookWGAU = XLSX.utils.book_new();
const workbookWSAH = XLSX.utils.book_new();
const workbookWCHA = XLSX.utils.book_new();
const workbookWDEL = XLSX.utils.book_new();
const workbookEORI = XLSX.utils.book_new();

let response1;
let response2;

// Dynamic date calculation with Kolkata timezone
const kolkataTimezone = 'Asia/Kolkata';
const today = moment().tz(kolkataTimezone);

let firstDate, current_date, dateDifference;

// Always use previous month start and end dates
const previousMonth = moment().tz(kolkataTimezone).subtract(1, 'month');
firstDate = previousMonth.startOf('month').format('YYYY-MM-DD');
current_date = previousMonth.endOf('month').format('YYYY-MM-DD');

dateDifference = moment(current_date).diff(moment(firstDate), 'days') + 1;

console.log('First Date (Previous Month Start):', firstDate);
console.log('Current Date (Previous Month End):', current_date);
console.log('Date Difference:', dateDifference);
console.log('Previous Month:', previousMonth.format('MMMM YYYY'));

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
                // result = result.filter(i => i['Current Status'] == 'Verified' && i['Branch Code']);
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

async function querydb() {
    let getMpduBaseSheetData = await getDataFromGoogleSheets(spreadsheetId, 'BaseSheetCall');
    console.log("MPDU BASE DATA SHEET ::", workbookData['All Device'].length);

    if (getMpduBaseSheetData) {
        response1 = await pool.query(`select * from display_data_table where custom_date = '${current_date}'`);
        response2 = await pool.query(`select * from display_data_table where custom_date >= '${firstDate}' and custom_date <= '${current_date}'`);

        console.log("Data From display_data_table ::", response2.rows.length);
    }
}

function createBranchSpecificWorkbooks(name, dataArray) {
    let wbhoarray = []
    let wnagarray = []
    let wvizarray = []
    let whydarray = []
    let wblrarray = []
    let wchearray = []
    let wjprarray = []
    let wmumarray = []
    let wpunarray = []
    let wlucarray = []
    let weuparray = []
    let wcalarray = []
    let wgauarray = []
    let wsaharray = []
    let wchaarray = []
    let wdelarray = []
    let eoriarray = []

    wbhoarray.push(dataArray[0])
    wnagarray.push(dataArray[0])
    wvizarray.push(dataArray[0])
    whydarray.push(dataArray[0])
    wblrarray.push(dataArray[0])
    wchearray.push(dataArray[0])
    wjprarray.push(dataArray[0])
    wmumarray.push(dataArray[0])
    wpunarray.push(dataArray[0])
    wlucarray.push(dataArray[0])
    weuparray.push(dataArray[0])
    wcalarray.push(dataArray[0])
    wgauarray.push(dataArray[0])
    wsaharray.push(dataArray[0])
    wchaarray.push(dataArray[0])
    wdelarray.push(dataArray[0])
    eoriarray.push(dataArray[0])


    wbhoarray.push(dataArray[1])
    wnagarray.push(dataArray[2])
    wvizarray.push(dataArray[3])
    whydarray.push(dataArray[4])
    wblrarray.push(dataArray[5])
    wchearray.push(dataArray[6])
    wjprarray.push(dataArray[7])
    wmumarray.push(dataArray[8])
    wpunarray.push(dataArray[9])
    wlucarray.push(dataArray[10])
    weuparray.push(dataArray[11])
    eoriarray.push(dataArray[12])
    wcalarray.push(dataArray[13])
    wgauarray.push(dataArray[14])
    wsaharray.push(dataArray[15])
    wchaarray.push(dataArray[16])
    wdelarray.push(dataArray[17])
    const wbhoarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wbhoarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWBHO, wbhoarray_array_to_sheet, name);
    //
    const wnagarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wnagarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWNAG, wnagarray_array_to_sheet, name);
    //
    const wvizarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wvizarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWVIZ, wvizarray_array_to_sheet, name);
    //
    const whydarray_array_to_sheet = XLSX.utils.aoa_to_sheet(whydarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWHYD, whydarray_array_to_sheet, name);
    //
    const wblrarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wblrarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWBLR, wblrarray_array_to_sheet, name);
    //
    const wchearray_array_to_sheet = XLSX.utils.aoa_to_sheet(wchearray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWCHE, wchearray_array_to_sheet, name);
    //
    const wjprarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wjprarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWJPR, wjprarray_array_to_sheet, name);
    //
    const wmumarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wmumarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWMUM, wmumarray_array_to_sheet, name);
    //
    const wpunarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wpunarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWPUN, wpunarray_array_to_sheet, name);
    //
    const wlucarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wlucarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWLUC, wlucarray_array_to_sheet, name);
    //
    const weuparray_array_to_sheet = XLSX.utils.aoa_to_sheet(weuparray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWEUP, weuparray_array_to_sheet, name);
    //
    const wcalarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wcalarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWCAL, wcalarray_array_to_sheet, name);
    //
    const wgauarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wgauarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWGAU, wgauarray_array_to_sheet, name);
    //
    const wsaharray_array_to_sheet = XLSX.utils.aoa_to_sheet(wsaharray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWSAH, wsaharray_array_to_sheet, name);

    //
    const wchaarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wchaarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWCHA, wchaarray_array_to_sheet, name);
    //
    const wdelarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wdelarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWDEL, wdelarray_array_to_sheet, name);

    const eoriarray_array_to_sheet = XLSX.utils.aoa_to_sheet(eoriarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookEORI, eoriarray_array_to_sheet, name);
}

function createDetailedBranchWorkbooks(name, dataArray) {
    let wbhoarray = []
    let wnagarray = []
    let wvizarray = []
    let whydarray = []
    let wblrarray = []
    let wchearray = []
    let wjprarray = []
    let wmumarray = []
    let wpunarray = []
    let wlucarray = []
    let weuparray = []
    let wcalarray = []
    let wgauarray = []
    let wsaharray = []
    let wchaarray = []
    let wdelarray = []
    let eoriarray = []

    wbhoarray.push(dataArray[0])
    wnagarray.push(dataArray[0])
    wvizarray.push(dataArray[0])
    whydarray.push(dataArray[0])
    wblrarray.push(dataArray[0])
    wchearray.push(dataArray[0])
    wjprarray.push(dataArray[0])
    wmumarray.push(dataArray[0])
    wpunarray.push(dataArray[0])
    wlucarray.push(dataArray[0])
    weuparray.push(dataArray[0])
    wcalarray.push(dataArray[0])
    wgauarray.push(dataArray[0])
    wsaharray.push(dataArray[0])
    wchaarray.push(dataArray[0])
    wdelarray.push(dataArray[0])
    eoriarray.push(dataArray[0])

    for (let index = 1; index < dataArray.length; index++) {
        for (let n = 0; n < dataArray[index].length; n++) {
            if (dataArray[index][n] === 'WBHO') {
                wbhoarray.push(dataArray[index])
                break
            }
            else if (dataArray[index][n] === 'WNAG') {
                wnagarray.push(dataArray[index])
                break
            }
            else if (dataArray[index][n] === 'EVIZ') {
                wvizarray.push(dataArray[index])
                break
            }
            else if (dataArray[index][n] === 'SHYD') {
                whydarray.push(dataArray[index])
                break
            }
            else if (dataArray[index][n] === 'SBLR') {
                wblrarray.push(dataArray[index])
                break
            }
            else if (dataArray[index][n] === 'SCHE') {
                wchearray.push(dataArray[index])
                break
            }
            else if (dataArray[index][n] === 'NJPR') {
                wjprarray.push(dataArray[index])
                break
            }
            else if (dataArray[index][n] === 'WMUM') {
                wmumarray.push(dataArray[index])
                break
            }
            else if (dataArray[index][n] === 'WPUN') {
                wpunarray.push(dataArray[index])
                break
            }
            else if (dataArray[index][n] === 'NLUC') {
                wlucarray.push(dataArray[index])
                break
            }
            else if (dataArray[index][n] === 'NEUP') {
                weuparray.push(dataArray[index])
                break
            }
            else if (dataArray[index][n] === 'ECAL') {
                wcalarray.push(dataArray[index])
                break
            }
            else if (dataArray[index][n] === 'EGAU') {
                wgauarray.push(dataArray[index])
                break
            }
            else if (dataArray[index][n] === 'NSAH') {
                wsaharray.push(dataArray[index])
                break
            }
            else if (dataArray[index][n] === 'NCHA') {
                wchaarray.push(dataArray[index])
                break
            }
            else if (dataArray[index][n] === 'NDEL') {
                wdelarray.push(dataArray[index])
                break
            }
            else if (dataArray[index][n] === 'EORI') {
                eoriarray.push(dataArray[index])
                break
            }
        }

    }
    //
    const wbhoarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wbhoarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWBHO, wbhoarray_array_to_sheet, name);
    //
    const wnagarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wnagarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWNAG, wnagarray_array_to_sheet, name);
    //
    const wvizarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wvizarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWVIZ, wvizarray_array_to_sheet, name);
    //
    const whydarray_array_to_sheet = XLSX.utils.aoa_to_sheet(whydarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWHYD, whydarray_array_to_sheet, name);
    //
    const wblrarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wblrarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWBLR, wblrarray_array_to_sheet, name);
    //
    const wchearray_array_to_sheet = XLSX.utils.aoa_to_sheet(wchearray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWCHE, wchearray_array_to_sheet, name);
    //
    const wjprarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wjprarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWJPR, wjprarray_array_to_sheet, name);
    //
    const wmumarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wmumarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWMUM, wmumarray_array_to_sheet, name);
    //
    const wpunarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wpunarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWPUN, wpunarray_array_to_sheet, name);
    //
    const wlucarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wlucarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWLUC, wlucarray_array_to_sheet, name);
    //
    const weuparray_array_to_sheet = XLSX.utils.aoa_to_sheet(weuparray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWEUP, weuparray_array_to_sheet, name);
    //
    const wcalarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wcalarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWCAL, wcalarray_array_to_sheet, name);
    //
    const wgauarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wgauarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWGAU, wgauarray_array_to_sheet, name);
    //
    const wsaharray_array_to_sheet = XLSX.utils.aoa_to_sheet(wsaharray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWSAH, wsaharray_array_to_sheet, name);

    //
    const wchaarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wchaarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWCHA, wchaarray_array_to_sheet, name);
    //
    const wdelarray_array_to_sheet = XLSX.utils.aoa_to_sheet(wdelarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookWDEL, wdelarray_array_to_sheet, name);

    const eoriarray_array_to_sheet = XLSX.utils.aoa_to_sheet(eoriarray);
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbookEORI, eoriarray_array_to_sheet, name);
}

function getUniqueValues(arr) {
    // Create a Set to store unique values
    const uniqueSet = new Set();

    // Iterate through the array and add each element to the Set
    for (const element of arr) {
        uniqueSet.add(element);
    }

    // Convert the Set back to an array (if needed)
    const uniqueArray = [...uniqueSet];

    return uniqueArray;
}

function aggregateStatusCountsByBranch(data) {
    const statusFilters = [
        "Never Installed",
        "Never Installed & Physical Damage",
        "UNKNOWN",
        "Verified & Currently Installed",
        "Verified & Outlet Closed",
        "Verified & Physical Damage",
        "Verified & Temp Closed"
    ];

    const result = {};

    data.forEach(item => {
        const branchCode = item["Branch Code"];
        const branchName = item["City"] || "";
        const status = item["Current Status"];

        if (!statusFilters.includes(status)) return;

        // For WNAG and EVIZ, use branch + city to separate groups
        const key = (branchCode === "WNAG" || branchCode === "EVIZ") ? `${branchCode}_${branchName}` : branchCode;

        if (!result[key]) {
            result[key] = {
                "Branch Code": branchCode,
                "City": branchName
            };
            statusFilters.forEach(s => result[key][s] = 0);
        }

        result[key][status]++;
    });

    // Convert 0s to "" for clean output
    Object.values(result).forEach(entry => {
        statusFilters.forEach(status => {
            if (entry[status] === 0) {
                entry[status] = "";
            }
        });
    });

    // Group WNAG by city under finalResult["WNAG"]
    // Group EVIZ by city under finalResult["EVIZ"]
    const finalResult = {};

    Object.values(result).forEach(entry => {
        const branch = entry["Branch Code"];

        if (branch === "WNAG") {
            if (!finalResult["WNAG"]) finalResult["WNAG"] = [];
            finalResult["WNAG"].push(entry);
        } else if (branch === "EVIZ") {
            if (!finalResult["EVIZ"]) finalResult["EVIZ"] = [];
            finalResult["EVIZ"].push(entry);
        } else {
            finalResult[branch] = entry;
        }
    });

    return finalResult;
}

function getDeviceDataByBranchAndCity(workbookData, sheetKey, branchName, cityName) {
    const batchData = workbookData[sheetKey] || [];

    // First, filter entries with matching branch
    const branchMatches = batchData.filter(entry =>
        (entry.Branch || '').trim().toLowerCase() === branchName.trim().toLowerCase()
    );

    // Then, find entry with matching city within that branch
    const matchedEntry = branchMatches.find(entry =>
        (entry.City || '').trim().toLowerCase() === cityName.trim().toLowerCase()
    );


    if (matchedEntry) {
        return {
            dispatchCount: Number(matchedEntry['Total Devices']) || 0,
            installedCount: Number(matchedEntry['Verified & Working']) || 0
        };
    } else {
        return {
            dispatchCount: 0,
            installedCount: 0
        };
    }
}

async function generateMTDSummaryReport() {
    let runtimeCount = 0;
    let percentageActiveDevices = 0;
    let activeDevicesCount = [];
    let runtimeInHours = 0;
    let poorcount = 0;
    let goodcount = 0;
    let verygoodcount = 0;
    let criticalcount = 0;
    let statusData;

    const dataArray1 = [[
        'Branch code',
        'Branch Name',
        'Phase 1 Dispatch',
        'Phase 2 Dispatch ',
        'Installed in Phase 1',
        'Installed in Phase 2',
        'Balance from Phase 1',
        'Balance from Phase 2',
        'Never Installed',
        'Never Installed & Physical Damage',
        'UNKNOWN',
        'Verified & Currently Installed',
        'Verified & Outlet Closed',
        'Verified & Physical Damage',
        '⁠Verified & Temp Closed',
        '% Active Devices',
        'Average Run Time Per Day'
    ]];

    function calculateDeviceScore(activedays, averageDailyRuntime, operationDays) {
        let t1 = (activedays / operationDays) * 10;
        let t2 = Math.floor(averageDailyRuntime / 8 * 10);
        if (t2 > 10) t2 = 10;
        const score = Math.ceil(t1 + t2);

        if (score >= 17) verygoodcount++;
        else if (score >= 12) goodcount++;
        else if (score >= 7) poorcount++;
        else criticalcount++;
    }

    let dim = [];
    let branchArray = [];
    const branchArr = [
        { branch: 'NLUC', 'branch-name': 'Lucknow' },
        { branch: 'WMUM', 'branch-name': 'Mumbai' },
        { branch: 'NEUP', 'branch-name': 'Varanasi' },
        { branch: 'SHYD', 'branch-name': 'Hyderabad' },
        { branch: 'NJPR', 'branch-name': 'Jaipur' },
        { branch: 'EVIZ', 'branch-name': 'Vizag' },
        { branch: 'EVIZ', 'branch-name': 'Visakhapatnam' },
        { branch: 'WNAG', 'branch-name': 'Nagpur' },
        { branch: 'WNAG', 'branch-name': 'Raipur' },
        { branch: 'WBHO', 'branch-name': 'Bhopal' },
        { branch: 'SBLR', 'branch-name': 'Bangalore' },
        { branch: 'SCHE', 'branch-name': 'Chennai' },
        { branch: 'WPUN', 'branch-name': 'Pune' },
        { branch: 'EORI', 'branch-name': 'Orissa' },
        { branch: 'ECAL', 'branch-name': 'Kolkata' },
        { branch: 'EGAU', 'branch-name': 'Guwahati' },
        { branch: 'NSAH', 'branch-name': 'Saharanpur' },
        { branch: 'NCHA', 'branch-name': 'Chandigarh' },
        { branch: 'NDEL', 'branch-name': 'Delhi' },
        { branch: 'SKAR', 'branch-name': 'Mangalore' },
        { branch: 'SCOI', 'branch-name': 'Coimbatore' },
        { branch: 'WAHM', 'branch-name': 'Ahmedabad' },
        { branch: 'SERN', 'branch-name': 'Ernakulam' },
        { branch: '', 'branch-name': 'Miscellaneous' }
    ];

    let excelArrayData = workbookData['All Device'];
    let excelArrayDataVerifiedDeviceList = excelArrayData.filter(d => d['Current Status'] == 'Verified & Currently Installed');
    let statusCountsByBranch = aggregateStatusCountsByBranch(excelArrayData);
    excelArrayDataVerifiedDeviceList.forEach(e => branchArray.push(e['Branch Code']));
    branchArray = getUniqueValues(branchArray);

    branchArr.forEach(branch => {
        const filteredbranch = excelArrayDataVerifiedDeviceList.filter(d => d['Branch Code'] === branch.branch);

        filteredbranch.forEach(deviceIdElement => {
            const deviceInDatabase = response2.rows.find(d => d.display_name == deviceIdElement['Techworks ID'] && d.display_count > 0);
            if (deviceInDatabase !== undefined) activeDevicesCount.push(deviceInDatabase);
        });

        filteredbranch.forEach(deviceIdElement => {
            const deviceInDatabase = response2.rows.filter(d => d.display_name == deviceIdElement['Techworks ID']);
            let operationDays = dateDifference;
            let activedays = 0;
            let runtimeAddition = 0;

            for (let index = 0; index < deviceInDatabase.length; index++) {
                if (Number(deviceInDatabase[index].display_count) > 0) {
                    activedays++;
                    runtimeAddition += Number(deviceInDatabase[index].display_count);
                }
            }

            let averageDailyRuntime = ((runtimeAddition / 4) / operationDays);
            calculateDeviceScore(activedays, averageDailyRuntime, operationDays);

            deviceInDatabase.forEach(dbElement => {
                if (dbElement !== undefined) {
                    runtimeCount += Number(dbElement.display_count);
                }
            });
        });

        percentageActiveDevices = Math.round((activeDevicesCount.length / filteredbranch.length) * 100);
        if (!isFinite(percentageActiveDevices)) percentageActiveDevices = 0;

        runtimeInHours = ((runtimeCount * 15) / 60) / filteredbranch.length;
        if (!isFinite(runtimeInHours)) runtimeInHours = 0;

        // this code is used for finding count filterwise
        if (branch.branch === 'WNAG' || branch.branch === 'EVIZ') {
            // WNAG or EVIZ is an array of city objects
            statusData = (statusCountsByBranch[branch.branch] || []).find(
                entry => (entry.City || '').trim().toLowerCase() === branch['branch-name'].trim().toLowerCase()
            );
        } else {
            statusData = statusCountsByBranch[branch.branch];
        }


        // Helper to get status count or 0 if missing
        const getStatus = (key) => statusData ? statusData[key] : "";

        const batch1Record = getDeviceDataByBranchAndCity(workbookData, 'Batch-1', branch.branch, branch['branch-name']);
        const batch2Record = getDeviceDataByBranchAndCity(workbookData, 'Batch-2', branch.branch, branch['branch-name']);

        // console.log("Batch 1", branch['branch'], branch['branch-name'], batch1Record);
        // console.log("Batch 2", branch['branch'], branch['branch-name'], batch2Record);


        dim.push(branch['branch']);
        dim.push(branch['branch-name']);
        dim.push(batch1Record?.dispatchCount);
        dim.push(batch2Record?.dispatchCount);
        dim.push(batch1Record?.installedCount);
        dim.push(batch2Record?.installedCount);
        dim.push(0);
        dim.push(0);
        if (branch['branch-name'] != 'Miscellaneous') {
            dim.push(getStatus('Never Installed'));
            dim.push(getStatus('Never Installed & Physical Damage'));
            dim.push(getStatus('UNKNOWN'));
            dim.push(getStatus('Verified & Currently Installed'));
            dim.push(getStatus('Verified & Outlet Closed'));
            dim.push(getStatus('Verified & Physical Damage'));
            dim.push(getStatus('Verified & Temp Closed'));
            dim.push(percentageActiveDevices + "%");
            dim.push((runtimeInHours / dateDifference).toFixed(2));
        } else {
            dim.push('');
            dim.push('');
            dim.push('');
            dim.push('');
            dim.push('');
            dim.push('');
            dim.push('');
            dim.push('');
            dim.push('');
        }

        dataArray1.push(dim);

        // Reset for next branch
        dim = [];
        activeDevicesCount = [];
        percentageActiveDevices = 0;
        runtimeCount = 0;
        criticalcount = 0;
        verygoodcount = 0;
        goodcount = 0;
        poorcount = 0;
    });

    // Add total row
    dim.push("Total");
    dim.push('');
    dim.push('');
    dim.push('');
    dim.push('');
    dim.push('');
    dim.push('');
    dataArray1.push(dim);

    const array_to_sheet = XLSX.utils.aoa_to_sheet(dataArray1);
    XLSX.utils.book_append_sheet(workbook, array_to_sheet, 'MTD Summary Report');
    createBranchSpecificWorkbooks('MTD Summary Report', dataArray1);
}

async function generateMTDDetailedReport() {
    let runtime = 0;
    let activeDays = 0;
    let averageDailyRuntime = 0;
    let metricEfficiency = 0;
    let city = '';
    let language = '';
    let outletName = '';
    let outletAddress = ''
    let outletContactNumber = ''
    let branchPoc = '';
    let dateOfInspection = '';
    let branchCode = '';
    let bucket = '';
    let percentageActiveDays = 0;
    let isactive = 0;
    let rundayscoring = 0
    let timescore = 0
    let t1 = 0
    let t2 = 0
    let cummalativescore;
    let cummalativerating;
    const dataArray = [[
        'Display Id', 'Date', 'City', 'Language', 'Outlet Name', 'Outlet Address', 'Outlet Contact Number', 'Branch Code', 'Branch POC', 'Date Of Inspection', 'Operation Days', 'Active Days', '% Active Days', 'Runtime', 'Average Daily Runtime', "Active", "Run Day Scoring", "Time score", "Run days Scoring (Max 10) No of days Active/ No of Total days in Month *10 Max Score", "Run Time Scoring (Max 10) Avg. No of Hour Active/ Avg. 8 Hours Run *10 Max Score", "Cummalative Score", "Cummalative Rating"
    ]];
    let dim = [];

    let excelArrayData = workbookData['All Device'].filter(d => d['Current Status'] == 'Verified & Currently Installed');

    // console.log(excelArrayData);
    let count = 0;
    excelArrayData.forEach(element => {
        // console.log(response.rows);
        const idDataFromDatabase = response2.rows.filter(d => {
            return d.display_name == element['Techworks ID'];
        });
        count++
        idDataFromDatabase.forEach(deviceElement => {
            // console.log(deviceElement.display_name, element['Techworks ID'], count);
            runtime += Number(deviceElement.display_count);
            if (Number(deviceElement.display_count) > 0) {
                activeDays++
                isactive = 1
            }
        });
        averageDailyRuntime = (runtime / 4) / dateDifference;

        if (idDataFromDatabase === undefined) {
            city = '';
            language = ''
            outletName = '';
            outletAddress = '';
            outletContactNumber = '';
            branchCode = '';
            branchPoc = '';
            dateOfInspection = '';
        } else {
            city = element['City'];
            language = element['Langauge'];
            outletName = element['Outlet Name'];
            outletAddress = element['Outlet Address'];
            outletContactNumber = element['Outlet Contact Number'];
            branchCode = element['Branch Code'];
            branchPoc = element['Branch POC'];
            dateOfInspection = element['Date of Inspection'];
        }

        metricEfficiency =
            (activeDays / dateDifference) *
            (averageDailyRuntime / 8) *
            100;


        if (Math.round(metricEfficiency) >= 80) {
            bucket = "Above 80";
        }
        if (
            Math.round(metricEfficiency) < 80 &&
            Math.round(metricEfficiency) >= 50
        ) {
            bucket = "Below 80";
        }
        if (Math.round(metricEfficiency) < 50 && Math.round(metricEfficiency) > 0) {
            bucket = "Below 50";
        }
        if (Math.round(metricEfficiency) === 0) {
            bucket = "Zero";
        }


        percentageActiveDays = (activeDays / dateDifference) * 100


        //remarks 
        if (Math.round(percentageActiveDays) >= 80 && Math.round(averageDailyRuntime) >= 8) {
            remarks = 'Outstanding'
        } if ((Math.round(percentageActiveDays) >= 70 && Math.round(percentageActiveDays) < 80) && Math.round(averageDailyRuntime) >= 8) {
            remarks = 'Good'
        } if (Math.round(percentageActiveDays) <= 70 && Math.round(averageDailyRuntime) > 0) {
            remarks = 'Average'
        } if (Math.round(percentageActiveDays) > 70 && Math.round(averageDailyRuntime) < 8) {
            remarks = 'Satisfactory'
        } if (Math.round(percentageActiveDays) <= 70 && Math.round(averageDailyRuntime) === 0) {
            remarks = 'Poor'
        }

        let operationDays = dateDifference
        timescore = averageDailyRuntime / 8
        rundayscoring = (activeDays / operationDays)
        t1 = (activeDays / operationDays) * 10
        if (Math.floor(averageDailyRuntime / 8 * 10) > 10) {  //doubt
            t2 = 10
        } else {
            t2 = Math.floor(averageDailyRuntime / 8 * 10)
        }
        cummalativescore = Math.floor(t1 + t2)

        if (cummalativescore >= 17 && cummalativescore <= 20) {
            cummalativerating = 'Very Good'
        } else if (cummalativescore >= 12 && cummalativescore <= 16) {
            cummalativerating = 'Good'
        } else if (cummalativescore >= 7 && cummalativescore <= 11) {
            cummalativerating = 'Poor'
        } else if (cummalativescore < 7) {
            cummalativerating = 'Critical'
        }

        if (idDataFromDatabase[0]?.display_name) {
            dim.push(idDataFromDatabase[0]?.display_name);
            dim.push(current_date);
            dim.push(city);
            dim.push(language);
            dim.push(outletName);
            dim.push(outletAddress);
            dim.push(outletContactNumber);
            dim.push(branchCode);
            dim.push(branchPoc);
            dim.push(dateOfInspection);
            dim.push(dateDifference);
            dim.push(activeDays);
            dim.push(((activeDays / dateDifference) * 100).toFixed(2) + "%");
            dim.push(runtime / 4);   //   15min/60
            dim.push(averageDailyRuntime.toFixed(2)); //average daily runtime
            // dim.push(metricEfficiency.toFixed(2) + "%");
            // dim.push(bucket);
            // dim.push(remarks);
            dim.push(isactive)
            dim.push((rundayscoring * 100).toFixed(2) + '%')
            dim.push((timescore * 100).toFixed(2) + '%')
            dim.push(t1.toFixed(2))
            dim.push(t2)
            dim.push(cummalativescore)
            dim.push(cummalativerating)

            dataArray.push(dim)
        }

        activeDays = 0;
        runtime = 0;
        averageDailyRuntime = 0;
        metricEfficiency = 0;
        percentageActiveDays = 0;
        bucket = ''
        remarks = '';
        isactive = 0;
        rundayscoring = 0
        timescore = 0
        t1 = 0
        t2 = 0
        cummalativescore = 0
        cummalativerating = null
        dim = [];
    });



    const array_to_sheet = XLSX.utils.aoa_to_sheet(dataArray);

    XLSX.utils.book_append_sheet(workbook, array_to_sheet, 'MTD Detailed Report');
    // XLSX.writeFile(workbook, 'combined_data.xlsx');
    createDetailedBranchWorkbooks('MTD Detailed Report', dataArray)
}

async function sendReportEmail(params) {
    try {
        // send mail with defined transport object
        const info = await transporter.sendMail({
            from: 'reports@techworks.co.in', // sender address
            // to:'pratik@techworks.co.in',
            // to:'hitesh.kumar@techworks.co.in',
            cc: "mark.thomas.k@gmail.com, rohanwork2002@gmail.com, shayan.p.sadique@gmail.com, dhruv@techworks.co.in",
            to: "reports@techworks.co.in, sumit.gupta@techworks.co.in, sandip@techworks.co.in, pratik@techworks.co.in", // list of receivers 
            subject: "MPDU REPORT TILL" + current_date, // Subject line
           html: `<h6>Please find the attachment.</h6>
            <p>&nbsp;</p>
            <table style="width:450px; font-size: 10pt; font-family: Verdana, sans-serif; background: transparent !important;"
                border="0" cellspacing="0" cellpadding="0">
                <tbody>
                    <tr>
                        <td style="width: 200px; font-size: 10pt; font-family: Verdana, sans-serif; vertical-align: top;"
                            valign="top">
                            <p style="margin-bottom: 18px; padding: 0px;"><span
                                    style="font-size: 12pt; font-family: Verdana, sans-serif; color: #183884; font-weight: bold;">Techworks
                                    Reports<br /></span><span
                                    style="font-family: Verdana, sans-serif; font-size: 9pt; color: #183884;">DS Techworks
                                    Solutions</span></p>
                            <p style="margin-top: 0px; margin-bottom: 18px; padding: 0px;"><a
                                    href="http://www.vcard.techworksworld.com/techworks_reports/" target="_blank"><img
                                        style="width: 120px; height: auto; border: 0;"
                                        src="https://raw.githubusercontent.com/tw-designer/tw-emp-qr-links/main/qr_techworks_reports.png"
                                        width="120" border="0" /></a></p>
                            <p
                                style="margin-bottom: 0px; padding: 0px; font-family: Verdana, sans-serif; font-size: 9pt; line-height: 12pt;">
                                <a style="color: #e25422; text-decoration: none; font-weight: bold;"
                                    href="http://www.techworksworld.com" rel="noopener"><span
                                        style="text-decoration: none; font-size: 9pt; line-height: 12pt; color: #e25422; font-family: Verdana, sans-serif; font-weight: bold;">www.techworksworld.com</span></a>
                            </p>
                        </td>
                        <td style="width: 10px; min-width: 10px; border-right: 1px solid #e25422;">&nbsp;</td>
                        <td style="width: 10px; min-width: 10px;">&nbsp;</td>
                        <td style="width: 250px; font-size: 10pt; color: #444444; font-family: Verdana, sans-serif; vertical-align: top;"
                            valign="top">
                            <p
                                style="font-family: Verdana, sans-serif; padding: 0px; font-size: 9pt; line-height: 14pt; margin-bottom: 14px;">
                                <span style="font-family: Verdana, sans-serif; font-size: 9pt; line-height: 14pt;"><span
                                        style="font-size: 9pt; line-height: 13pt; color: #262626;"><strong>E: </strong></span><a
                                        style="font-size: 9pt; color: #262626; text-decoration: none;"
                                        href="mailto:reports@techworks.co.in"><span
                                            style="text-decoration: none; font-size: 9pt; line-height: 14pt; color: #262626; font-family: Verdana, sans-serif;">reports@techworks.co.in</span></a><span><br /></span></span><span><span
                                        style="font-size: 9pt; color: #262626;"><strong>T:</strong></span><span
                                        style="font-size: 9pt; color: #262626;">(+91)
                                        8920131195</span><span><br /></span></span><span><span
                                        style="font-size: 9pt; color: #262626;"><strong>A:</strong></span><span
                                        style="font-size: 9pt; color: #262626;"> O-7, 2nd Floor Lajpat Nagar-II, </span><span
                                        style="color: #262626;">New Delhi-110024, India</span></span></p>
                            <p style="margin-bottom: 0px; padding: 0px;"><span><a
                                        href="https://www.facebook.com/TechworksSolutionsPvtLtd/" rel="noopener"><img
                                            style="border: 0; height: 22px; width: 22px;"
                                            src="https://www.mail-signatures.com/signature-generator/img/templates/inclusive/fb.png"
                                            width="22" border="0" /></a>&nbsp;</span><span><a
                                        href="https://www.linkedin.com/company/ds-techworks-solutions-pvt-ltd/" rel="noopener"><img
                                            style="border: 0; height: 22px; width: 22px;"
                                            src="https://www.mail-signatures.com/signature-generator/img/templates/inclusive/ln.png"
                                            width="22" border="0" /></a>&nbsp;</span><span><a href="https://twitter.com/techworks14"
                                        rel="noopener"><img style="border: 0; height: 22px; width: 22px;"
                                            src="https://www.mail-signatures.com/signature-generator/img/templates/inclusive/tt.png"
                                            width="22" border="0" /></a>&nbsp;</span><span><a
                                        href="https://www.youtube.com/@TechworksDigitalSolutions" rel="noopener"><img
                                            style="border: 0; height: 22px; width: 22px;"
                                            src="https://www.mail-signatures.com/signature-generator/img/templates/inclusive/yt.png"
                                            width="22" border="0" /></a>&nbsp;</span><span><a
                                        href="https://www.instagram.com/techworks140/" rel="noopener"><img
                                            style="border: 0; height: 22px; width: 22px;"
                                            src="https://www.mail-signatures.com/signature-generator/img/templates/inclusive/it.png"
                                            width="22" border="0" /></a></span></p>
                        </td>
                    </tr>
                    <tr style="width: 420px;">
                        <td style="padding-top: 14px;" colspan="4"><a href="https://techworksworld.com/" rel="noopener"><img
                                    style="width: 420px; height: auto; border: 0;" src="https://i.imgur.com/QoPxSPy.png" width="420"
                                    border="0" /></a></td>
                    </tr>
                    <tr>
                        <td style="padding-top: 14px; text-align: justify;" colspan="4">
                            <table
                                style="width: 420px; font-size: 10pt; font-family: Verdana, sans-serif; background: transparent !important;"
                                border="0" cellspacing="0" cellpadding="0">
                                <tbody>
                                    <tr>
                                        <td style="font-size: 8pt; color: #b2b2b2; line-height: 9pt; text-align: justify;">The
                                            content of this email is confidential and intended for the recipient specified in
                                            message only. It is strictly forbidden to share any part of this message with any third
                                            party,without a written consent of the sender. If you received this message by
                                            mistake,please reply to this message and follow with its deletion,so that we can ensure
                                            such a mistake does not occur in the future.</td>
                                    </tr>
                                </tbody>
                            </table>
                        </td>
                    </tr>
                </tbody>
            </table>`,
            attachments: [
                {   // file on disk as an attachment
                    filename: `MPDU REPORT TILL${current_date}.xlsx`,
                    path: `./reports/MPDU REPORT TILL${current_date}.xlsx` // stream this file
                },
                {   // file on disk as an attachment
                    filename: `MPDU REPORT TILL ECAL${current_date}.xlsx`,
                    path: `./reports/MPDU REPORT TILL ECAL${current_date}.xlsx` // stream this file
                },
                {   // file on disk as an attachment
                    filename: `MPDU REPORT TILL EGAU${current_date}.xlsx`,
                    path: `./reports/MPDU REPORT TILL EGAU${current_date}.xlsx` // stream this file
                },
                {   // file on disk as an attachment
                    filename: `MPDU REPORT TILL EVIZ${current_date}.xlsx`,
                    path: `./reports/MPDU REPORT TILL EVIZ${current_date}.xlsx` // stream this file
                },
                {   // file on disk as an attachment
                    filename: `MPDU REPORT TILL NCHA${current_date}.xlsx`,
                    path: `./reports/MPDU REPORT TILL NCHA${current_date}.xlsx` // stream this file
                },
                {   // file on disk as an attachment
                    filename: `MPDU REPORT TILL NDEL${current_date}.xlsx`,
                    path: `./reports/MPDU REPORT TILL NDEL${current_date}.xlsx` // stream this file
                },
                {   // file on disk as an attachment
                    filename: `MPDU REPORT TILL NEUP${current_date}.xlsx`,
                    path: `./reports/MPDU REPORT TILL NEUP${current_date}.xlsx` // stream this file
                },
                {   // file on disk as an attachment
                    filename: `MPDU REPORT TILL NJPR${current_date}.xlsx`,
                    path: `./reports/MPDU REPORT TILL NJPR${current_date}.xlsx` // stream this file
                },
                {   // file on disk as an attachment
                    filename: `MPDU REPORT TILL NLUC${current_date}.xlsx`,
                    path: `./reports/MPDU REPORT TILL NLUC${current_date}.xlsx` // stream this file
                },
                {   // file on disk as an attachment
                    filename: `MPDU REPORT TILL NSAH${current_date}.xlsx`,
                    path: `./reports/MPDU REPORT TILL NSAH${current_date}.xlsx` // stream this file
                },
                {   // file on disk as an attachment
                    filename: `MPDU REPORT TILL SBLR${current_date}.xlsx`,
                    path: `./reports/MPDU REPORT TILL SBLR${current_date}.xlsx` // stream this file
                },
                {   // file on disk as an attachment
                    filename: `MPDU REPORT TILL SCHE${current_date}.xlsx`,
                    path: `./reports/MPDU REPORT TILL SCHE${current_date}.xlsx` // stream this file
                },
                {   // file on disk as an attachment
                    filename: `MPDU REPORT TILL SHYD${current_date}.xlsx`,
                    path: `./reports/MPDU REPORT TILL SHYD${current_date}.xlsx` // stream this file
                },
                {   // file on disk as an attachment
                    filename: `MPDU REPORT TILL WBHO${current_date}.xlsx`,
                    path: `./reports/MPDU REPORT TILL WBHO${current_date}.xlsx` // stream this file
                },
                {   // file on disk as an attachment
                    filename: `MPDU REPORT TILL WMUM${current_date}.xlsx`,
                    path: `./reports/MPDU REPORT TILL WMUM${current_date}.xlsx` // stream this file
                },
                {   // file on disk as an attachment
                    filename: `MPDU REPORT TILL WNAG${current_date}.xlsx`,
                    path: `./reports/MPDU REPORT TILL WNAG${current_date}.xlsx` // stream this file
                },
                {   // file on disk as an attachment
                    filename: `MPDU REPORT TILL WPUN${current_date}.xlsx`,
                    path: `./reports/MPDU REPORT TILL WPUN${current_date}.xlsx` // stream this file
                },
                {   // file on disk as an attachment
                    filename: `MPDU REPORT TILL EORI${current_date}.xlsx`,
                    path: `./reports/MPDU REPORT TILL EORI${current_date}.xlsx` // stream this file
                },
            ]
        });

        console.log("Mail Send Succesfull.. ")

    } catch (error) {
        console.log(error);
    }
}

Promise.all([querydb()])
    .then(() => {
        Promise.all([
            generateMTDSummaryReport(),
            generateMTDDetailedReport()
        ])
            .then(() => {
                setTimeout(() => {
                    // Check if 'reports' folder exists, if not create it
                    if (!fs.existsSync(reportsFolderPath)) {
                        fs.mkdirSync(reportsFolderPath);
                    }
                    
                    const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
                    const fileName = 'MPDU REPORT TILL' + current_date + '.xlsx'
                    XLSX.writeFile(workbook, path.join(reportsFolderPath, fileName));

                    // const buffersWBHO = XLSX.write(workbookWBHO, { bookType: 'xlsx', type: 'buffer' });
                    // const fileNameWBHO = 'MPDU REPORT TILL' + ' WBHO' + current_date + '.xlsx'
                    // XLSX.writeFile(workbookWBHO, path.join(reportsFolderPath, fileNameWBHO));

                    // //
                    // const buffersNAG = XLSX.write(workbookWNAG, { bookType: 'xlsx', type: 'buffer' });
                    // const fileNameWNAG = 'MPDU REPORT TILL' + ' WNAG' + current_date + '.xlsx'
                    // XLSX.writeFile(workbookWNAG, path.join(reportsFolderPath, fileNameWNAG));

                    // //
                    // const buffersWVIZ = XLSX.write(workbookWVIZ, { bookType: 'xlsx', type: 'buffer' });
                    // const fileNameWVIZ = 'MPDU REPORT TILL' + ' EVIZ' + current_date + '.xlsx'
                    // XLSX.writeFile(workbookWVIZ, path.join(reportsFolderPath, fileNameWVIZ));

                    // //
                    // const buffersWHYD = XLSX.write(workbookWHYD, { bookType: 'xlsx', type: 'buffer' });
                    // const fileNameWHYD = 'MPDU REPORT TILL' + ' SHYD' + current_date + '.xlsx'
                    // XLSX.writeFile(workbookWHYD, path.join(reportsFolderPath, fileNameWHYD));


                    // //
                    // const buffersWBLR = XLSX.write(workbookWBLR, { bookType: 'xlsx', type: 'buffer' });
                    // const fileNameWBLR = 'MPDU REPORT TILL' + ' SBLR' + current_date + '.xlsx'
                    // XLSX.writeFile(workbookWBLR, path.join(reportsFolderPath, fileNameWBLR));


                    // //
                    // const buffersWCHE = XLSX.write(workbookWCHE, { bookType: 'xlsx', type: 'buffer' });
                    // const fileNameWCHE = 'MPDU REPORT TILL' + ' SCHE' + current_date + '.xlsx'
                    // XLSX.writeFile(workbookWCHE, path.join(reportsFolderPath, fileNameWCHE));


                    // //
                    // const buffersWJPR = XLSX.write(workbookWJPR, { bookType: 'xlsx', type: 'buffer' });
                    // const fileNameWJPR = 'MPDU REPORT TILL' + ' NJPR' + current_date + '.xlsx'
                    // XLSX.writeFile(workbookWJPR, path.join(reportsFolderPath, fileNameWJPR));


                    // //
                    // const buffersWMUM = XLSX.write(workbookWMUM, { bookType: 'xlsx', type: 'buffer' });
                    // const fileNameWMUM = 'MPDU REPORT TILL' + ' WMUM' + current_date + '.xlsx'
                    // XLSX.writeFile(workbookWMUM, path.join(reportsFolderPath, fileNameWMUM));


                    // //
                    // const buffersWPUN = XLSX.write(workbookWPUN, { bookType: 'xlsx', type: 'buffer' });
                    // const fileNameWPUN = 'MPDU REPORT TILL' + ' WPUN' + current_date + '.xlsx'
                    // XLSX.writeFile(workbookWPUN, path.join(reportsFolderPath, fileNameWPUN));


                    // //
                    // const buffersWLUC = XLSX.write(workbookWLUC, { bookType: 'xlsx', type: 'buffer' });
                    // const fileNameWLUC = 'MPDU REPORT TILL' + ' NLUC' + current_date + '.xlsx'
                    // XLSX.writeFile(workbookWLUC, path.join(reportsFolderPath, fileNameWLUC));


                    // //
                    // const buffersWEUP = XLSX.write(workbookWEUP, { bookType: 'xlsx', type: 'buffer' });
                    // const fileNameWEUP = 'MPDU REPORT TILL' + ' NEUP' + current_date + '.xlsx'
                    // XLSX.writeFile(workbookWEUP, path.join(reportsFolderPath, fileNameWEUP));


                    // //
                    // const buffersWCAL = XLSX.write(workbookWCAL, { bookType: 'xlsx', type: 'buffer' });
                    // const fileNameWCAL = 'MPDU REPORT TILL' + ' ECAL' + current_date + '.xlsx'
                    // XLSX.writeFile(workbookWCAL, path.join(reportsFolderPath, fileNameWCAL));


                    // //
                    // const buffersWGAU = XLSX.write(workbookWGAU, { bookType: 'xlsx', type: 'buffer' });
                    // const fileNameWGAU = 'MPDU REPORT TILL' + ' EGAU' + current_date + '.xlsx'
                    // XLSX.writeFile(workbookWGAU, path.join(reportsFolderPath, fileNameWGAU));


                    // //
                    // const buffersWSAH = XLSX.write(workbookWSAH, { bookType: 'xlsx', type: 'buffer' });
                    // const fileNameWSAH = 'MPDU REPORT TILL' + ' NSAH' + current_date + '.xlsx'
                    // XLSX.writeFile(workbookWSAH, path.join(reportsFolderPath, fileNameWSAH));


                    // //
                    // const buffersWCHA = XLSX.write(workbookWCHA, { bookType: 'xlsx', type: 'buffer' });
                    // const fileNameWCHA = 'MPDU REPORT TILL' + ' NCHA' + current_date + '.xlsx'
                    // XLSX.writeFile(workbookWCHA, path.join(reportsFolderPath, fileNameWCHA));


                    // //
                    // const buffersWDEL = XLSX.write(workbookWDEL, { bookType: 'xlsx', type: 'buffer' });
                    // const fileNameWDEL = 'MPDU REPORT TILL' + ' NDEL' + current_date + '.xlsx'
                    // XLSX.writeFile(workbookWDEL, path.join(reportsFolderPath, fileNameWDEL));

                    // const buffersEORI = XLSX.write(workbookEORI, { bookType: 'xlsx', type: 'buffer' });
                    // const fileNameEORI = 'MPDU REPORT TILL' + ' EORI' + current_date + '.xlsx'
                    // XLSX.writeFile(workbookEORI, path.join(reportsFolderPath, fileNameEORI));
                }, 3000);

                setTimeout(() => {
                    // sendReportEmail();
                    // reportDeliveryNDEL();
                }, 10000);
            })
            .catch((error) => {
                console.error("An error occurred:", error);
            })
    })
