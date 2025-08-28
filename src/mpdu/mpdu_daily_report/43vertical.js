const nodemailer = require("nodemailer");
const { google } = require('googleapis');
const { Pool } = require('pg');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const CONSTANTS = require("./constants");
const { getCredentialsPath } = require("../../utility/pathUtils");


let accountList = [
    {
        "user": "pratik@techworks.co.in",
        "pass": "Ew^KvQkh"
    },
    {
        "user": "chirag.p@techworks.co.in",
        "pass": "Byzzy1-jucton-gogkeq"
    },
    {
        "user": "Bharti.singh@techworks.co.in",
        "pass": "gymruc-saKpu8-purnoc"
    },
    {
        "user": "Aaditya@techworks.co.in",
        "pass": "r19yb5E*syU1kRFa"
    },
    {
        "user": "hitesh.kumar@techworks.co.in",
        "pass": "4VqvS&RY*ZFnqaU1"
    },
]

const transporter1 = nodemailer.createTransport({
    host: "smtp.dreamhost.com",
    port: 465,
    secure: true,
    auth: {
        user: accountList[0].user,
        pass: accountList[0].pass
    }
});

const transporter2 = nodemailer.createTransport({
    host: "smtp.dreamhost.com",
    port: 465,
    secure: true,
    auth: {
        user: accountList[1].user,
        pass: accountList[1].pass
    }
});

const transporter3 = nodemailer.createTransport({
    host: "smtp.dreamhost.com",
    port: 465,
    secure: true,
    auth: {
        user: accountList[2].user,
        pass: accountList[2].pass
    }
});

const transporter4 = nodemailer.createTransport({
    host: "smtp.dreamhost.com",
    port: 465,
    secure: true,
    auth: {
        user: accountList[3].user,
        pass: accountList[3].pass
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
const workbooks = {};

let response1;
let response2;

// Define allowed branches
const ALLOWED_BRANCHES = [
    'WMUM', 'ECAL', 'NDEL', 'NCHA', 'WPUN', 'NJPR', 'SBLR'
]

// Email configurations for different branches
const EMAIL_CONFIG = {
    'DEFAULT': {
        to: "mark.thomas.k@gmail.com, rohanwork2002@gmail.com, shayan.p.sadique@gmail.com",
        cc: "dhruv@techworks.co.in, sumit.gupta@techworks.co.in, sandip@techworks.co.in, pratik@techworks.co.in, rusum@techworks.co.in, hitesh.kumar@techworks.co.in",
        subject: "43 VERTICAL REPORT TILL",
        transporterName: "transporter1",
        emailCount: 8
    },
    'ECAL': {
        to: "Rajdeep.Datta@itc.in, Dipannita.Tiwary@itc.in",
        cc: "sumit.gupta@techworks.co.in, sandip@techworks.co.in, pratik@techworks.co.in, rusum@techworks.co.in",
        subject: "ECAL 43 VERTICAL REPORT TILL",
        transporterName: "transporter1",
        emailCount: 8
    },
    'NCHA': {
        to: "v.bhardwaj@itc.in",
        cc: "sumit.gupta@techworks.co.in, sandip@techworks.co.in, pratik@techworks.co.in, rusum@techworks.co.in",
        subject: "NCHA 43 VERTICAL REPORT TILL",
        transporterName: "transporter1",
        emailCount: 7
    },
    'NDEL': {
        to: "Amit.Srivastava@itc.in",
        cc: "sumit.gupta@techworks.co.in, sandip@techworks.co.in, pratik@techworks.co.in, rusum@techworks.co.in",
        subject: "NDEL 43 VERTICAL REPORT TILL",
        transporterName: "transporter2",
        emailCount: 7
    },
    'NJPR': {
        to: "Himanshu.Tanwar@itc.in, AkshatSingh.Ranavat@itc.in",
        cc: "sumit.gupta@techworks.co.in, sandip@techworks.co.in, pratik@techworks.co.in, rusum@techworks.co.in",
        subject: "NJPR 43 VERTICAL REPORT TILL",
        transporterName: "transporter2",
        emailCount: 8
    },
    'SBLR': {
        to: "Shreyas.K@itc.in, MohammedYahya.Zaid@itc.in",
        cc: "sumit.gupta@techworks.co.in, sandip@techworks.co.in, pratik@techworks.co.in, rusum@techworks.co.in",
        subject: "SBLR 43 VERTICAL REPORT TILL",
        transporterName: "transporter2",
        emailCount: 7
    },
    'WMUM': {
        to: "Bibhu.Priyadarshi@itc.in, mohammed.glasswala@itc.in, rishabsunil.jain@itc.in, Akash.Sagar@itc.in",
        cc: "sumit.gupta@techworks.co.in, sandip@techworks.co.in, pratik@techworks.co.in, rusum@techworks.co.in",
        subject: "WMUM 43 VERTICAL REPORT TILL",
        transporterName: "transporter2",
        emailCount: 7
    },
    'WPUN': {
        to: "kritika.mahajan@itc.in, Danish.Sayyed@itc.in",
        cc: "sumit.gupta@techworks.co.in, sandip@techworks.co.in, pratik@techworks.co.in, rusum@techworks.co.in",
        subject: "WPUN 43 VERTICAL REPORT TILL",
        transporterName: "transporter2",
        emailCount: 7
    }
}


async function getDataFromGoogleSheets(sheetID, reference) {
    const accessGoogleSheet = async () => {
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

    const getAllWorkbookNames = async (sheets) => {
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

    const getWorkbookWiseData = async (sheets, sheetNames) => {
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
                result = result.filter(i => i['Current Status'] == 'Verified & Working' && i['Branch Code']);
                workbookData[sheetName] = result;
            }
        } catch (error) {
            console.error("Error fetching data for sheets:", error);
            throw error;
        }

        return true;
    };

    try {
        let sheets = await accessGoogleSheet();
        let sheetNames = await getAllWorkbookNames(sheets);
        let getWorkbookRes = await getWorkbookWiseData(sheets, ['43 Inch Vertical']);
        return getWorkbookRes;
    } catch (error) {
        console.error("Error during Google Sheets data retrieval:", error);
    }
}

async function querydb() {
    console.log("\n");
    console.log("===============================================");
    console.log(`             DATE :-  ${current_date}          `);
    console.log("===============================================");
    let get43VerticalBaseSheetData = await getDataFromGoogleSheets(spreadsheetId, 'BaseSheetCall');
    console.log("43-VERTICAL BASE DATA SHEET ::", workbookData['43 Inch Vertical'].length);

    if (get43VerticalBaseSheetData) {
        response1 = (await pool.query(`select * from display_data_table where custom_date = '${current_date}'`)).rows;
        response2 = (await pool.query(`select * from display_data_table where custom_date >= '${firstDate}' and custom_date <= '${current_date}'`)).rows;

        console.log("Data From display_data_table ::", response1.length, response2.length);
    }
}

function common(name, dataArray) {
    // Skip the header row (index 0) and total row (last row)
    for (let i = 1; i < dataArray.length - 1; i++) {
        const branchCode = dataArray[i][0];  // Branch code is the first column

        // Only process if it's in our allowed branches list
        if (ALLOWED_BRANCHES.includes(branchCode)) {
            // Create a new workbook for this branch if it doesn't exist
            if (!workbooks[branchCode]) {
                workbooks[branchCode] = XLSX.utils.book_new();
            }

            // Create array with headers and branch data
            const branchArray = [
                dataArray[0],  // Headers
                dataArray[i]   // Branch specific data
            ];

            // Convert array to sheet
            const worksheet = XLSX.utils.aoa_to_sheet(branchArray);

            // Add worksheet to workbook
            XLSX.utils.book_append_sheet(workbooks[branchCode], worksheet, name);
        }
    }
}

// Add function to get file path for reports folder
function getFilePath(filename) {
    const reportsDir = path.join(__dirname, 'reports');

    // Ensure reports directory exists
    if (!fs.existsSync(reportsDir)) {
        fs.mkdirSync(reportsDir, { recursive: true });
    }

    return path.join(reportsDir, filename);
}

function detailedreport(name, dataArray) {
    // Create a map to store arrays for each branch
    const branchArrays = {};

    // Initialize arrays for all allowed branches with headers
    ALLOWED_BRANCHES.forEach(branch => {
        branchArrays[branch] = [dataArray[0]];
    });

    // Process data rows
    for (let i = 1; i < dataArray.length; i++) {
        for (let j = 0; j < dataArray[i].length; j++) {
            const value = dataArray[i][j];
            // Check if the value is in our allowed branches
            if (ALLOWED_BRANCHES.includes(value)) {
                branchArrays[value].push(dataArray[i]);
                break;
            }
        }
    }

    // Convert arrays to sheets and add to workbooks
    ALLOWED_BRANCHES.forEach(branchCode => {
        // Only create sheets for branches that have data
        if (branchArrays[branchCode].length > 1) { // More than just headers
            // Convert array to sheet
            const worksheet = XLSX.utils.aoa_to_sheet(branchArrays[branchCode]);

            // Create workbook if it doesn't exist
            if (!workbooks[branchCode]) {
                workbooks[branchCode] = XLSX.utils.book_new();
            }

            // Add worksheet to workbook
            XLSX.utils.book_append_sheet(workbooks[branchCode], worksheet, name);
        }
    });
}
const dispatched = {
    'WMUM': 25,
    'NJPR': 6,
    'SBLR': 35,
    'WPUN': 5,
    'ECAL': 4,
    'NCHA': 5,
    'NDEL': 70
};
const today = new Date();
const oneDayMilliseconds = 24 * 60 * 60 * 1000; // Number of milliseconds in one day
const oneDayBefore = new Date(today.getTime() - oneDayMilliseconds);

const year = oneDayBefore.getFullYear();
const month = String(oneDayBefore.getMonth() + 1).padStart(2, "0");
const day = String(oneDayBefore.getDate()).padStart(2, "0");

const firstDate = ` ${year}-${month}-01`;
// console.log(firstDate);
// const current_date = `${year}-${month}-${day}`;
const current_date = `2025-08-19`;
const dateDifference = new Date(current_date).getDate() + 1 - new Date(firstDate).getDate();
// console.log(current_date);


function findUniqueValues(arr) {
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

async function dailySummaryReport() {
    // console.log("\n");
    // console.log("===============================================");
    // console.log("       DAILY SUMMARY REPORT MAKING START       ");
    // console.log("===============================================");

    let runtimeCount = 0;
    let percentageActiveDevices = 0;
    let activeDevicesCount = [];
    let runtimeInHours = 0;
    let totalDevices = 0;
    let totalActiveDevices = 0
    let sumProductAverageRuntimeTemp = 0;
    let sumProductAverageRuntime = 0;
    let totaldispatched = 0
    const dataArray = [[
        'Branch', 'Total Devices', 'Active Devices', 'Inactive Devices', '% Active Devices', 'Average Runtime', 'Optimal Runtime', 'Runtime Efficacy', "Dispatched Device"
    ]];
    let dim = [];
    let branchArray = [];

    //getting all branch
    workbookData['43 Inch Vertical'].forEach(excelBranchElement => {
        branchArray.push(excelBranchElement[CONSTANTS.BRANCH]);
    });

    // filter unqiue branch
    branchArray = findUniqueValues(branchArray);

    branchArray.forEach((element, index) => {
        const filteredbranch = workbookData['43 Inch Vertical'].filter(d => d[CONSTANTS.BRANCH] === element); //filter excel row acc. to brancharray 

        filteredbranch.forEach(deviceIdElement => {
            const deviceInDatabase = response1.find(d => d.display_name === deviceIdElement[CONSTANTS.TECHWORKS_ID]);

            if (deviceInDatabase?.display_count !== undefined) {
                runtimeCount += Number(deviceInDatabase?.display_count);
            }

            if (deviceInDatabase?.display_count > 0) {
                activeDevicesCount.push(deviceInDatabase);
            }
        });

        percentageActiveDevices = Math.round((Number(activeDevicesCount.length) / Number(filteredbranch.length)) * 100);
        if (!isFinite(percentageActiveDevices) || isNaN(percentageActiveDevices)) {
            percentageActiveDevices = 0;
        }

        runtimeInHours = ((runtimeCount * 15) / 60) / filteredbranch.length;
        if (!isFinite(runtimeInHours) || isNaN(runtimeInHours)) {
            runtimeInHours = 0;
        }


        dim.push(element);
        dim.push(filteredbranch.length);
        dim.push(activeDevicesCount.length);
        dim.push(Number(filteredbranch.length) - Number(activeDevicesCount.length));
        dim.push(percentageActiveDevices + "%");
        dim.push(runtimeInHours.toFixed(2));
        dim.push(8);
        dim.push(((runtimeInHours / 8) * 100).toFixed(2) + "%")
        dim.push(dispatched[element])
        dataArray.push(dim);

        // reset all vaiables
        dim = []
        activeDevicesCount = [];
        percentageActiveDevices = 0;
        runtimeCount = 0;
    });
    for (let i = 1; i < dataArray.length; i++) {
        const element = dataArray[i];
        totalDevices += element[1];
        totalActiveDevices += element[2];
        sumProductAverageRuntimeTemp = Number(element[5] * element[2]);
        sumProductAverageRuntime += sumProductAverageRuntimeTemp;
        totaldispatched += element[8]
    }

    dim.push("Total");
    dim.push(totalDevices);
    dim.push(totalActiveDevices);
    dim.push(totalDevices - totalActiveDevices);
    dim.push(Math.round((totalActiveDevices / totalDevices) * 100) + "%");
    dim.push((sumProductAverageRuntime / totalActiveDevices).toFixed(2))
    dim.push(8);
    dim.push(((((Math.trunc(sumProductAverageRuntime / totalActiveDevices)) / 8) * 100)).toFixed(2) + "%");
    dim.push(totaldispatched)
    dataArray.push(dim);

    const array_to_sheet = XLSX.utils.aoa_to_sheet(dataArray);
    XLSX.utils.book_append_sheet(workbook, array_to_sheet, 'Daily Summary Report');
    common('Daily Summary Report', dataArray);
}

async function mtdSummaryReport() {
    // console.log("\n");
    // console.log("===============================================");
    // console.log("        MTD SUMMARY REPORT MAKING START        ");
    // console.log("===============================================");

    let runtimeCount = 0;
    let percentageActiveDevices = 0;
    let activeDevicesCount = [];
    let runtimeInHours = 0;
    let totalDevices = 0;
    let totalActiveDevices = 0
    let sumProductAverageRuntimeTemp = 0;
    let sumProductAverageRuntime = 0;
    let poorcount = 0
    let goodcount = 0
    let verygoodcount = 0
    let criticalcount = 0
    let totalcriticalcount = 0
    let totalpoorcount = 0
    let totalgoodcount = 0
    let totalverygoodcount = 0
    let totaldispatched = 0
    const dataArray = [[
        'Branch', 'Total Devices', 'Active Devices', 'Inactive Devices', '% Active Devices', 'Average Runtime', 'Optimal Runtime', 'Runtime Efficacy', "Critical", "Poor", "Good", "Very Good", "Dispatched Device"
    ]];

    function cal(activedays, averageDailyRuntime, operationDays) {
        let cummalativescore = 0;

        t1 = (activedays / operationDays) * 10
        if (Math.floor(averageDailyRuntime / 8 * 10) > 10) {
            t2 = 10;
        } else {
            t2 = Math.floor(averageDailyRuntime / 8 * 10);
        }
        cummalativescore = Math.ceil(t1 + t2);
        if (cummalativescore >= 17 && cummalativescore <= 20) {
            // cummalativerating = 'Very Good'
            verygoodcount++
        } else if (cummalativescore >= 12 && cummalativescore <= 16) {
            // cummalativerating = 'Good'
            goodcount++
        } else if (cummalativescore >= 7 && cummalativescore <= 11) {
            // cummalativerating = 'Poor'
            poorcount++
        } else if (cummalativescore < 7) {

            criticalcount++
        }
    }
    let dim = [];
    let branchArray = [];

    workbookData['43 Inch Vertical'].forEach(excelBranchElement => {
        branchArray.push(excelBranchElement[CONSTANTS.BRANCH]);
    });
    branchArray = findUniqueValues(branchArray);

    branchArray.forEach((element, index) => {
        const filteredbranch = workbookData['43 Inch Vertical'].filter(d => d[CONSTANTS.BRANCH] === element);

        filteredbranch.forEach(deviceIdElement => {
            const deviceInDatabase = response2.find(d => d.display_name == deviceIdElement[CONSTANTS.TECHWORKS_ID] && d.display_count > 0);

            if (deviceInDatabase !== undefined) {
                activeDevicesCount.push(deviceInDatabase);
            }
        });

        filteredbranch.forEach(deviceIdElement => {
            const deviceInDatabase = response2.filter(d => d.display_name == deviceIdElement[CONSTANTS.TECHWORKS_ID]);

            let operationDays = dateDifference
            let activedays = 0
            let runtimeAddition = 0
            for (let index = 0; index < deviceInDatabase.length; index++) {
                if (Number(deviceInDatabase[index].display_count) > 0) {
                    activedays++
                    runtimeAddition += Number(deviceInDatabase[index].display_count)
                }
            }
            let averageDailyRuntime = ((Number(runtimeAddition) / 4) / operationDays)
            cal(activedays, averageDailyRuntime, operationDays)

            deviceInDatabase.forEach(dbElement => {
                if (dbElement !== undefined) {
                    runtimeCount += Number(dbElement.display_count)
                }
            });

        });

        percentageActiveDevices = Math.round((Number(activeDevicesCount.length) / Number(filteredbranch.length)) * 100);
        if (!isFinite(percentageActiveDevices) || isNaN(percentageActiveDevices)) {
            percentageActiveDevices = 0;
        }
        runtimeInHours = ((runtimeCount * 15) / 60) / filteredbranch.length;

        if (!isFinite(runtimeInHours) || isNaN(runtimeInHours)) {
            runtimeInHours = 0;
        }

        dim.push(element);
        dim.push(filteredbranch.length);
        dim.push(activeDevicesCount.length);
        dim.push(Number(filteredbranch.length) - Number(activeDevicesCount.length));
        dim.push(percentageActiveDevices + "%");
        dim.push((runtimeInHours / dateDifference).toFixed(2));
        dim.push(8);
        dim.push((((runtimeInHours / dateDifference) / 8) * 100).toFixed(2) + "%")
        dim.push(criticalcount)
        dim.push(poorcount)
        dim.push(goodcount)
        dim.push(verygoodcount)
        dim.push(dispatched[element])
        dataArray.push(dim);

        // reset all variables
        dim = []
        activeDevicesCount = [];
        percentageActiveDevices = 0;
        runtimeCount = 0;
        criticalcount = 0
        verygoodcount = 0
        goodcount = 0
        poorcount = 0

    });
    for (let i = 1; i < dataArray.length; i++) {
        const element = dataArray[i];
        totalDevices += element[1];
        totalActiveDevices += element[2];
        sumProductAverageRuntimeTemp = Number(element[5] * element[2]);
        sumProductAverageRuntime += sumProductAverageRuntimeTemp;
        totalverygoodcount += element[8]
        totalgoodcount += element[9]
        totalpoorcount += element[10]
        totalcriticalcount += element[11]
        totaldispatched += element[12]
    }
    dim.push("Total");
    dim.push(totalDevices);
    dim.push(totalActiveDevices);
    dim.push(totalDevices - totalActiveDevices);
    dim.push(Math.round((totalActiveDevices / totalDevices) * 100) + "%");
    dim.push((sumProductAverageRuntime / totalActiveDevices).toFixed(2))
    dim.push(8);
    dim.push(((((Math.trunc(sumProductAverageRuntime / totalActiveDevices)) / 8) * 100)).toFixed(2) + "%");
    dim.push(totalverygoodcount)
    dim.push(totalgoodcount)
    dim.push(totalpoorcount)
    dim.push(totalcriticalcount)
    dim.push(totaldispatched)
    dataArray.push(dim);

    const array_to_sheet = XLSX.utils.aoa_to_sheet(dataArray);
    XLSX.utils.book_append_sheet(workbook, array_to_sheet, 'MTD Summary Report');
    common('MTD Summary Report', dataArray)
}

async function dailyReportDetailed() {
    // console.log("\n");
    // console.log("===============================================");
    // console.log("       DAILY DETAIL REPORT MAKING START        ");
    // console.log("===============================================");

    let city = '';
    let language = '';
    let outletName = '';
    let outletAddress = ''
    let outletContactNumber = ''
    let branchCode = '';
    let branchPoc = '';
    let aeName = '';
    let dateOfInspection = '';
    const dataArray = [[
        'Display Name', 'Date', 'City', 'Language', 'Outlet Name', 'Outlet Address', 'Outlet Contact Number', 'Branch Code', 'Branch POC', 'AE Name', 'Date Of Inspection', 'Runtime'
    ]];
    let dim = [];

    workbookData['43 Inch Vertical'].forEach((element, index) => {
        const idDataFromDatabase = response1.find(d => d.display_name.replace(/\s*(\(new\)|\t)\s*/gi, '') == element[CONSTANTS.TECHWORKS_ID]);

        if (idDataFromDatabase === undefined) {
            city = '';
            language = ''
            outletName = '';
            outletAddress = '';
            outletContactNumber = '';
            branchCode = '';
            branchPoc = '';
            dateOfInspection = '';
            aeName = '';
        } else {
            city = element[CONSTANTS.CITY];
            language = element[CONSTANTS.LANGAUAGE];
            outletName = element[CONSTANTS.OUTLET_NAME];
            outletAddress = element[CONSTANTS.OUTLET_ADDRESS];
            outletContactNumber = element[CONSTANTS.OUTLET_CONTACT_NUM];
            branchCode = element[CONSTANTS.BRANCH];
            branchPoc = element[CONSTANTS.BRANCH_POC_NAME];
            dateOfInspection = element[CONSTANTS.DATE_OF_INSPECTION];
            aeName = element[CONSTANTS.AE_NAME];
        }


        dim.push(idDataFromDatabase?.display_name);
        dim.push((idDataFromDatabase?.custom_date)?.toString().substring(4, 15));
        dim.push(city);
        dim.push(language);
        dim.push(outletName);
        dim.push(outletAddress);
        dim.push(outletContactNumber);
        dim.push(branchCode);
        dim.push(branchPoc);
        dim.push(aeName);
        dim.push(dateOfInspection);
        dim.push(idDataFromDatabase?.display_count == 0 ? 0 : ((idDataFromDatabase?.display_count * 15) / 60).toFixed(2));
        dataArray.push(dim);

        // rest all variables
        city = '';
        language = ''
        outletName = '';
        outletAddress = '';
        outletContactNumber = '';
        branchCode = '';
        branchPoc = '';
        aeName = '';
        dateOfInspection = '';
        dim = [];
    });


    const array_to_sheet = XLSX.utils.aoa_to_sheet(dataArray);
    XLSX.utils.book_append_sheet(workbook, array_to_sheet, 'Daily Detailed Report');
    detailedreport('Daily Detailed Report', dataArray)
}

async function mtdReportDetailed() {
    // console.log("\n");
    // console.log("===============================================");
    // console.log("        MTD REPORT DETAILED MAKING START       ");
    // console.log("===============================================");

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
    let simcardnumber = '';
    let simCardProvider = '';
    let branchCode = '';
    let aeName = '';
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
        'Display Name', 'Date', 'City', 'Language', 'Outlet Name', 'Outlet Address', 'Outlet Contact Number', 'Branch Code', 'Branch POC', 'AE Name', 'Date Of Inspection', 'Sim Card Number', 'Sim Card Provider', 'Operation Days', 'Active Days', '% Active Days', 'Runtime', 'Average Daily Runtime', 'New Metric Efficiency', 'Bucket', 'Remarks', "Active", "Run Day Scoring", "Time score", "Run days Scoring (Max 10) No of days Active/ No of Total days in Month *10 Max Score", "Run Time Scoring (Max 10) Avg. No of Hour Active/ Avg. 8 Hours Run *10 Max Score", "Cummalative Score", "Cummalative Rating"
    ]];
    let dim = [];

    workbookData['43 Inch Vertical'].forEach(element => {
        const idDataFromDatabase = response2.filter(d => d.display_name.replace(/\s*(\(new\)|\t)\s*/gi, '') == element[CONSTANTS.TECHWORKS_ID]);
        const checkDeviceActiveOrNot = response1.filter(d => d.display_name.replace(/\s*(\(new\)|\t)\s*/gi, '') == element[CONSTANTS.TECHWORKS_ID] && Number(d.display_count) > 0);

        if (checkDeviceActiveOrNot.length > 0) {
            isactive = 1;
        } else {
            isactive = 0;
        }

        idDataFromDatabase.forEach(deviceElement => {
            runtime += Number(deviceElement.display_count);
            if (Number(deviceElement.display_count) > 0) {
                activeDays++
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
            simcardnumber = '';
            simCardProvider = '';
            aeName = '';
        } else {
            city = element[CONSTANTS.CITY];
            language = element[CONSTANTS.LANGAUAGE];
            outletName = element[CONSTANTS.OUTLET_NAME];
            outletAddress = element[CONSTANTS.OUTLET_ADDRESS];
            outletContactNumber = element[CONSTANTS.OUTLET_CONTACT_NUM];
            branchCode = element[CONSTANTS.BRANCH];
            branchPoc = element[CONSTANTS.BRANCH_POC_NAME];
            dateOfInspection = element[CONSTANTS.DATE_OF_INSPECTION];
            simcardnumber = element[CONSTANTS.SIM_CARD_NUM];
            simCardProvider = element[CONSTANTS.SIM_CARD_PROVIDER];
            aeName = element[CONSTANTS.AE_NAME];
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
        if (Math.floor(averageDailyRuntime / 8 * 10) > 10) {
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

        dim.push(idDataFromDatabase[0]?.display_name);
        dim.push(current_date);
        dim.push(city);
        dim.push(language);
        dim.push(outletName);
        dim.push(outletAddress);
        dim.push(outletContactNumber);
        dim.push(branchCode);
        dim.push(branchPoc);
        dim.push(aeName);
        dim.push(dateOfInspection);
        dim.push(simcardnumber);
        dim.push(simCardProvider);
        dim.push(dateDifference);
        dim.push(activeDays);
        dim.push(((activeDays / dateDifference) * 100).toFixed(2) + "%");
        dim.push(runtime / 4);   //   15min/60
        dim.push(averageDailyRuntime.toFixed(2)); //average daily runtime
        dim.push(metricEfficiency.toFixed(2) + "%");
        dim.push(bucket);
        dim.push(remarks);
        dim.push(isactive)
        dim.push((rundayscoring * 100).toFixed(2) + '%')
        dim.push((timescore * 100).toFixed(2) + '%')
        dim.push(t1.toFixed(2))
        dim.push(t2)
        dim.push(cummalativescore)
        dim.push(cummalativerating)
        dataArray.push(dim)

        // reset all variables
        activeDays = 0;
        runtime = 0;
        averageDailyRuntime = 0;
        metricEfficiency = 0;
        percentageActiveDays = 0;
        bucket = ''
        remarks = '';
        city = '';
        language = '';
        outletName = '';
        outletAddress = '';
        outletContactNumber = '';
        branchCode = '';
        branchPoc = '';
        aeName = '';
        dateOfInspection = '';
        simcardnumber = '';
        simCardProvider = '';
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
    detailedreport('MTD Detailed Report', dataArray)
}

async function dailyReportOnlineOffline() {
    // console.log("\n");
    // console.log("===============================================");
    // console.log("       43 VERTICAL DAILY ONLINE/OFFLINE REPORT MAKING START        ");
    // console.log("===============================================");

    // Generate time intervals for the entire day (15-minute intervals)
    const timeIntervals = [];
    const startTime = new Date(current_date + ' 00:00:00');
    const endTime = new Date(current_date + ' 23:59:59');

    for (let time = new Date(startTime); time <= endTime; time.setMinutes(time.getMinutes() + 15)) {
        const hours = time.getHours().toString().padStart(2, '0');
        const minutes = time.getMinutes().toString().padStart(2, '0');
        const ampm = hours >= 12 ? 'PM' : 'AM';
        const displayHours = hours > 12 ? (hours - 12).toString().padStart(2, '0') : (hours === '00' ? '12' : hours);
        timeIntervals.push(`${displayHours}:${minutes}${ampm}`);
    }

    // Create header row with only Device ID and time intervals
    const headerRow = ['Device ID'];

    // Add time interval columns
    timeIntervals.forEach(interval => {
        headerRow.push(interval);
    });

    // Add summary columns
    headerRow.push('Online Intervals');
    headerRow.push('Offline Intervals');
    headerRow.push('Online Percentage');

    const dataArray = [headerRow];

    // Process each device
    workbookData['43 Inch Vertical'].forEach((element) => {
        const deviceId = element[CONSTANTS.TECHWORKS_ID];

        // Create row for this device with only Device ID
        const deviceRow = [deviceId];

        // Get device data from database for current date
        const deviceData = response1.find(d => d.display_name.replace(/\s*(\(new\)|\t)\s*/gi, '') === deviceId);

        // Fill time interval columns based on hourly_status data
        let onlineCount = 0;

        timeIntervals.forEach((interval, index) => {
            let status = 0; // Default to offline (0)

            if (deviceData && deviceData.hourly_status) {
                try {
                    // Parse the hourly_status JSON
                    const hourlyStatus = typeof deviceData.hourly_status === 'string'
                        ? JSON.parse(deviceData.hourly_status)
                        : deviceData.hourly_status;

                    // Extract time part from interval (e.g., "12:00AM" -> "12:00")
                    const timePart = interval.replace(/[AP]M$/, '');
                    const ampm = interval.includes('AM') ? 'AM' : 'PM';

                    // Convert to 24-hour format for matching
                    let hour = parseInt(timePart.split(':')[0]);
                    const minute = timePart.split(':')[1];

                    if (ampm === 'PM' && hour !== 12) hour += 12;
                    if (ampm === 'AM' && hour === 12) hour = 0;

                    // Format hour and minute to match hourly_status format
                    const formattedHour = hour.toString().padStart(2, '0');
                    const formattedMinute = minute.padStart(2, '0');

                    // Create key to search in hourly_status
                    // Convert current_date from "2025-08-19" to "19-08-2025" format
                    const dateParts = current_date.split('-');
                    const formattedDate = `${dateParts[2]}-${dateParts[1]}-${dateParts[0]}`;
                    const searchKey = `${formattedDate}-${formattedHour}-${formattedMinute}-${ampm}`;



                    // Check if this key exists in hourly_status
                    if (hourlyStatus[searchKey]) {
                        const hourlyStatusValue = hourlyStatus[searchKey];
                        status = hourlyStatusValue === 'online' ? 1 : 0;
                    } else {
                        // If exact key not found, try to find the closest time entry
                        // This handles cases where data might be at different minute intervals
                        const availableKeys = Object.keys(hourlyStatus).filter(key =>
                            key.startsWith(formattedDate) &&
                            key.includes(`-${formattedHour}-`) &&
                            key.endsWith(`-${ampm}`)
                        );

                        if (availableKeys.length > 0) {
                            // Find the closest time entry by comparing minutes
                            let closestKey = availableKeys[0];
                            let minDifference = Math.abs(parseInt(formattedMinute) - parseInt(closestKey.split('-')[3]));

                            for (const key of availableKeys) {
                                const keyMinute = parseInt(key.split('-')[3]);
                                const difference = Math.abs(parseInt(formattedMinute) - keyMinute);
                                if (difference < minDifference) {
                                    minDifference = difference;
                                    closestKey = key;
                                }
                            }

                            const hourlyStatusValue = hourlyStatus[closestKey];
                            status = hourlyStatusValue === 'online' ? 1 : 0;


                        } else {
                            // If no keys found for this hour, try to find the closest available time from the entire day
                            const allAvailableKeys = Object.keys(hourlyStatus).filter(key => key.startsWith(formattedDate));

                            if (allAvailableKeys.length > 0) {
                                // Find the closest time by converting both to minutes since midnight
                                const targetMinutes = hour * 60 + parseInt(formattedMinute);
                                let closestKey = allAvailableKeys[0];
                                let minDifference = Infinity;

                                for (const key of allAvailableKeys) {
                                    const keyParts = key.split('-');
                                    const keyHour = parseInt(keyParts[3]);
                                    const keyMinute = parseInt(keyParts[4]);
                                    const keyAmpm = keyParts[5];

                                    // Convert key time to 24-hour format
                                    let keyHour24 = keyHour;
                                    if (keyAmpm === 'PM' && keyHour !== 12) keyHour24 += 12;
                                    if (keyAmpm === 'AM' && keyHour === 12) keyHour24 = 0;

                                    const keyMinutes = keyHour24 * 60 + keyMinute;
                                    const difference = Math.abs(targetMinutes - keyMinutes);

                                    if (difference < minDifference) {
                                        minDifference = difference;
                                        closestKey = key;
                                    }
                                }

                                const hourlyStatusValue = hourlyStatus[closestKey];
                                status = hourlyStatusValue === 'online' ? 1 : 0;


                            }
                        }
                    }
                } catch (error) {
                    console.error(`Error parsing hourly_status for device ${deviceId}:`, error);
                    status = 0; // Default to offline if parsing fails
                }
            }

            deviceRow.push(status);
            if (status === 1) onlineCount++;
        });

        // Add summary columns at the end
        const totalIntervals = timeIntervals.length;
        const offlineCount = totalIntervals - onlineCount;
        const onlinePercentage = totalIntervals > 0 ? ((onlineCount / totalIntervals) * 100).toFixed(2) : '0.00';

        deviceRow.push(onlineCount);
        deviceRow.push(offlineCount);
        deviceRow.push(onlinePercentage + '%');

        dataArray.push(deviceRow);
    });

    // Add totals row
    const totalDevices = dataArray.length - 1; // Subtract 1 for header row
    const totalOnlineIntervals = dataArray.slice(1).reduce((sum, row) => sum + parseInt(row[row.length - 3] || 0), 0);
    const totalOfflineIntervals = dataArray.slice(1).reduce((sum, row) => sum + parseInt(row[row.length - 2] || 0), 0);
    const totalIntervals = totalOnlineIntervals + totalOfflineIntervals;
    const overallOnlinePercentage = totalIntervals > 0 ? ((totalOnlineIntervals / totalIntervals) * 100).toFixed(2) : '0.00';

    const totalsRow = [
        'TOTAL',
        ...Array(timeIntervals.length).fill(''), // Empty cells for time intervals
        totalOnlineIntervals,
        totalOfflineIntervals,
        overallOnlinePercentage + '%'
    ];

    dataArray.push(totalsRow);

    const array_to_sheet = XLSX.utils.aoa_to_sheet(dataArray);
    XLSX.utils.book_append_sheet(workbook, array_to_sheet, 'Daily Online-Offline Report');
    detailedreport('Daily Online-Offline Report', dataArray);
}

async function sendBranchReport(branch = null) {
    try {
        // Determine email configuration based on branch
        const config = branch ? EMAIL_CONFIG[branch] || EMAIL_CONFIG.DEFAULT : EMAIL_CONFIG.DEFAULT;

        // Get the correct transporter based on config
        let currentTransporter;
        switch (config.transporterName) {
            case 'transporter1':
                currentTransporter = transporter1;
                break;
            case 'transporter2':
                currentTransporter = transporter2;
                break;
            case 'transporter3':
                currentTransporter = transporter3;
                break;
            case 'transporter4':
                currentTransporter = transporter4;
                break;
            default:
                currentTransporter = transporter1;
        }

        // Create attachments array
        const attachments = [];

        // For DEFAULT config or no specific branch, include all reports
        if (!branch || config === EMAIL_CONFIG.DEFAULT) {
            // Add main report
            attachments.push({
                filename: `43 VERTICAL REPORT TILL${current_date}.xlsx`,
                path: getFilePath(`43 VERTICAL REPORT TILL${current_date}.xlsx`)
            });

            // Add all branch reports
            ALLOWED_BRANCHES.forEach(branchCode => {
                attachments.push({
                    filename: `43 VERTICAL REPORT TILL ${branchCode}${current_date}.xlsx`,
                    path: getFilePath(`43 VERTICAL REPORT TILL ${branchCode}${current_date}.xlsx`)
                });
            });
        } else {
            // For specific branch, only include that branch's report
            attachments.push({
                filename: `43 VERTICAL REPORT TILL ${branch}${current_date}.xlsx`,
                path: getFilePath(`43 VERTICAL REPORT TILL ${branch}${current_date}.xlsx`)
            });
        }

        // Send email with configuration using the correct transporter
        const info = await currentTransporter.sendMail({
            from: 'reports@techworks.co.in',
            cc: config.cc,
            to: config.to,
            // to: 'hitesh.kumar@techworks.co.in',
            subject: config.subject + current_date,
            html: `<h6>Please find the attached 43 vertical report.</h6>`,
            attachments: attachments
        });

        console.log("\n");
        console.log("===============================================");
        console.log(`     MAIL SENT SUCCESSFULLY ${branch || 'ALL'}     `);
        console.log("===============================================");

        return info;
    } catch (error) {
        console.error(`Error sending email for ${branch || 'ALL'}:`, error);
        throw error;
    }
}

async function sendReports() {
    try {
        console.log("\n");
        console.log("===============================================");
        console.log("              STARTING REPORT DELIVERY         ");
        console.log("===============================================");

        // Send main report with all branches
        await sendBranchReport();

        // Group branches by transporter
        const transporterGroups = {
            transporter1: [],
            transporter2: [],
            transporter3: [],
            transporter4: []
        };

        // Sort branches into transporter groups
        ALLOWED_BRANCHES.forEach(branch => {
            const config = EMAIL_CONFIG[branch];
            if (config && config.transporterName) {
                transporterGroups[config.transporterName].push(branch);
            }
        });

        // Send reports for each transporter group sequentially
        for (const [transporter, branches] of Object.entries(transporterGroups)) {
            console.log(`\nSending emails for ${transporter}...`);

            for (const branch of branches) {
                try {
                    await sendBranchReport(branch);
                    // Add a small delay between emails to prevent rate limiting
                    await new Promise(resolve => setTimeout(resolve, 2000));
                } catch (error) {
                    console.error(`Failed to send report for ${branch}:`, error);
                    continue;
                }
            }
        }

        console.log("\n");
        console.log("===============================================");
        console.log("           ALL REPORTS SENT SUCCESSFULLY       ");
        console.log("===============================================");
    } catch (error) {
        console.error("Error in sendReports:", error);
        throw error;
    }
}

Promise.all([querydb()])
    .then(() => {
        Promise.all([
            dailySummaryReport(),
            mtdSummaryReport(),
            dailyReportDetailed(),
            mtdReportDetailed(),
            dailyReportOnlineOffline()
        ])
            .then(() => {
                setTimeout(() => {
                    // Write main workbook with proper path
                    const fileName = '43 VERTICAL REPORT TILL' + current_date + '.xlsx';
                    const mainFilePath = getFilePath(fileName);
                    XLSX.writeFile(workbook, mainFilePath);
                    console.log(`Main report saved: ${mainFilePath}`);

                    // Write individual branch workbooks with proper paths
                    Object.entries(workbooks).forEach(([branchCode, workbook]) => {
                        const branchFileName = `43 VERTICAL REPORT TILL ${branchCode}${current_date}.xlsx`;
                        const branchFilePath = getFilePath(branchFileName);
                        XLSX.writeFile(workbook, branchFilePath);
                        console.log(`Branch report saved: ${branchFilePath}`);
                    });

                    console.log("\n===============================================");
                    console.log("           ALL REPORTS SAVED SUCCESSFULLY       ");
                    console.log("===============================================");
                }, 3000);

                // setTimeout(() => {
                //     sendReports();
                // }, 10000);
            })
            .catch((error) => {
                console.error("An error occurred:", error);
            })
    })