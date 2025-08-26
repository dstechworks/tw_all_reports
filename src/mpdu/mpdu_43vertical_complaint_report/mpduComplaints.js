const logger = require('./mpduComlaintsLogger');
const moment = require('moment-timezone');
const { google } = require('googleapis');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const nodemailer = require("nodemailer");
const { getCredentialsPath } = require('../../utility/pathUtils');


let accountList = [
    {
        "user": "bharti.singh@techworks.co.in",
        "pass": "gymruc-saKpu8-purnoc"
    },
    {
        "user": "hitesh.kumar@techworks.co.in",
        "pass": "4VqvS&RY*ZFnqaU1"
    },
]

const transporter = nodemailer.createTransport({
    host: "smtp.dreamhost.com",
    port: 465,
    secure: true,
    auth: {
        user: accountList[1].user,
        pass: accountList[1].pass
    }
});

// Google Sheet ID
const spreadsheetId = "19D6cteChE9Plb-C7iDI6TbOD2jsWIpU6wgXRT4_2938";
// Get the current date in Asia/Kolkata timezone
const currentDate = moment().tz("Asia/Kolkata").toDate();
const currentFormattedDate = moment(currentDate).format('DD-MM-YYYY');
let workbookData = {};

// Define the folder path for saving the .xlsx files
const reportsFolderPath = path.join(__dirname, 'reports');
const fileName = 'MPDU_Complaints_Tracker_' + currentFormattedDate + '.xlsx'

function delay(milliseconds) {
    return new Promise(resolve => {
        setTimeout(resolve, milliseconds);
    });
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
                const columnsToRemove = [
                    'LAST ACTIVITY DATE',
                    'Problem Type',
                    'Number Of Days',
                    'Items To Be Replaced',
                    'Technician Visited',
                    'Expense',
                    'Handled By',
                    'Remarks 1',
                    'Tech Charge'
                ];

                // Filter out unwanted columns from headers
                const filteredHeaders = headers.filter(header => !columnsToRemove.includes(header));
                // console.log('Filtered Headers:', filteredHeaders);

                // Create objects with only the filtered headers
                const result = rows.map(row => {
                    const obj = {};
                    headers.forEach((key, index) => {
                        if (!columnsToRemove.includes(key)) {
                            obj[key] = row[index];
                        }
                    });
                    return obj;
                });

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
        let getWorkbookRes = await getWorkbookWiseData(sheets, sheetNames);
        return getWorkbookRes;
    } catch (error) {
        console.error("Error during Google Sheets data retrieval:", error);
    }
}

const createSeparateExcelFiles = async () => {
    let getBaseSheetData = await getDataFromGoogleSheets(spreadsheetId, 'BaseSheetCall');
    let baseDataSheet = workbookData["Master File"];
    console.log(baseDataSheet.length);

    if (baseDataSheet.length > 0) {
        // Filter data based on Date Of Complaint using Asia/Kolkata timezone
        // complaint 
        const startDate = moment.tz('01/04/2025', 'DD/MM/YYYY', 'Asia/Kolkata').startOf('day');
        const today = moment().tz('Asia/Kolkata').endOf('day');

        console.log('Filtering dates from:', startDate.format('DD/MM/YYYY'), 'to:', today.format('DD/MM/YYYY'));

        const filteredData = baseDataSheet.filter(row => {
            // Skip if Date Of Complaint is empty
            if (!row['Date Of Complaint']) {
                return false;
            }

            try {
                const complaintDate = moment.tz(row['Date Of Complaint'], 'DD/MM/YYYY', 'Asia/Kolkata');

                // Check if date is valid and within range
                if (!complaintDate.isValid()) {
                    console.log('Invalid date format:', row['Date Of Complaint']);
                    return false;
                }

                // Check if date is between April 1st and today
                const isInRange = complaintDate.isSameOrAfter(startDate) && complaintDate.isSameOrBefore(today);

                // if (isInRange) {
                //     console.log('Including date:', row['Date Of Complaint']);
                // }

                return isInRange;
            } catch (error) {
                console.log('Error processing date:', row['Date Of Complaint']);
                return false;
            }
        });

        // Sort the filtered data by Date Of Complaint (oldest to newest)
        const sortedData = filteredData.sort((a, b) => {
            const dateA = moment.tz(a['Date Of Complaint'], 'DD/MM/YYYY', 'Asia/Kolkata');
            const dateB = moment.tz(b['Date Of Complaint'], 'DD/MM/YYYY', 'Asia/Kolkata');
            return dateA - dateB;
        });

        console.log('Total records before filtering:', baseDataSheet.length);
        console.log('Records after date filtering:', sortedData.length);

        // Create a new workbook for each sheet
        const wb = xlsx.utils.book_new();
        const ws = xlsx.utils.json_to_sheet(sortedData);
        xlsx.utils.book_append_sheet(wb, ws, 'BaseSheetCall');
        const buffer = xlsx.write(wb, { bookType: 'xlsx', type: 'buffer' });
        xlsx.writeFile(wb, path.join(reportsFolderPath, fileName));

        await delay(4000);

        // Send email with attachment
        await sendEmail(fileName);
    }
}

async function sendEmail(fileName) {
    let fromEmail = 'reports@techworks.co.in';

    try {
        const mailOptions = {
            from: fromEmail,
            // to: 'hitesh.kumar@techworks.co.in',
            to: 'mark.thomas.k@gmail.com, rohanwork2002@gmail.com',
            cc: 'dhruv@techworks.co.in, rusum@techworks.co.in, sandip@techworks.co.in, bharti.singh@techworks.co.in, pratik@techworks.co.in, hitesh.kumar@techworks.co.in',
            subject: `MPDU Complaints Tracker - ${currentFormattedDate}`,
            html: `<h6>Please find MPDU Complaints Tracker Report.</h6>
            <p>&nbsp;</p>
            <table style="width: 420px; font-size: 10pt; font-family: Verdana, sans-serif; background: transparent !important;"
                    border="0" cellspacing="0" cellpadding="0">
                    <tbody>
                            <tr>
                                    <td style="width: 160px; font-size: 10pt; font-family: Verdana, sans-serif; vertical-align: top;"
                                            valign="top">
                                            <p style="margin-bottom: 18px; padding: 0px;"><span
                                                            style="font-size: 12pt; font-family: Verdana, sans-serif; color: #183884; font-weight: bold;"><a
                                                                    href="http://contact.techworksworld.com/" target="_"><img
                                                                            style="width: 120px; height: auto; border: 0;"
                                                                            src="https://i.imgur.com/Eie8F53.png" width="120"
                                                                            border="0" /></a></p>
                                            <p
                                                    style="margin-bottom: 0px; padding: 0px; font-family: Verdana, sans-serif; font-size: 9pt; line-height: 12pt;">
                                                    <a style="color: #e25422; text-decoration: none; font-weight: bold;"
                                                            href="http://www.techworksworld.com" target="_"><span
                                                                    style="text-decoration: none; font-size: 9pt; line-height: 12pt; color: #e25422; font-family: Verdana, sans-serif; font-weight: bold;">www.techworksworld.com</span></a>
                                            </p>
                                    </td>
                                    <td style="width: 30px; min-width: 30px; border-right: 1px solid #e25422;">&nbsp;</td>
                                    <td style="width: 30px; min-width: 30px;">&nbsp;</td>
                                    <td style="width: 200px; font-size: 10pt; color: #444444; font-family: Verdana, sans-serif; vertical-align: top;"
                                            valign="top">
                                            <p
                                                    style="font-family: Verdana, sans-serif; padding: 0px; font-size: 9pt; line-height: 14pt; margin-bottom: 14px;">
                                                    <span
                                                            style="font-family: Verdana, sans-serif; font-size: 9pt; line-height: 14pt;"><span
                                                                    style="font-size: 9pt; line-height: 13pt; color: #262626;"><strong>E:
                                                                    </strong></span><a
                                                                    style="font-size: 9pt; color: #262626; text-decoration: none;"
                                                                    href="mailto:support@techworks.co.in"><span
                                                                            style="text-decoration: none; font-size: 9pt; line-height: 14pt; color: #262626; font-family: Verdana, sans-serif;">support@techworks.co.in</span></a><span><br /></span></span><span><span
                                                                    style="font-size: 9pt; color: #262626;"><strong>T:</strong></span><span
                                                                    style="font-size: 9pt; color: #262626;"> (+91) 11 35007205</span><span><br /></span></span><span><span
                                                                    style="font-size: 9pt; color: #262626;"><strong>A:</strong></span><span
                                                                    style="font-size: 9pt; color: #262626;"> O-7,Second Floor Lajpat
                                                                    Nagar-II,<span>,</span></span><span style="color: #262626;">New
                                                                    Delhi-110024,India</span></span></p>
                                            <p style="margin-bottom: 0px; padding: 0px;"><span><a
                                                                    href="https://www.facebook.com/TechworksSolutionsPvtLtd/"
                                                                    rel="noopener"><img
                                                                            style="border: 0; height: 22px; width: 22px;"
                                                                            src="https://www.mail-signatures.com/signature-generator/img/templates/inclusive/fb.png"
                                                                            width="22" border="0" /></a>&nbsp;</span><span><a
                                                                    href="https://www.linkedin.com/company/ds-techworks-solutions-pvt-ltd/"
                                                                    rel="noopener"><img
                                                                            style="border: 0; height: 22px; width: 22px;"
                                                                            src="https://www.mail-signatures.com/signature-generator/img/templates/inclusive/ln.png"
                                                                            width="22" border="0" /></a>&nbsp;</span><span><a
                                                                    href="https://twitter.com/techworks14" rel="noopener"><img
                                                                            style="border: 0; height: 22px; width: 22px;"
                                                                            src="https://www.mail-signatures.com/signature-generator/img/templates/inclusive/tt.png"
                                                                            width="22" border="0" /></a>&nbsp;</span><span><a
                                                                    href="https://www.youtube.com/@TechworksDigitalSolutions"
                                                                    rel="noopener"><img
                                                                            style="border: 0; height: 22px; width: 22px;"
                                                                            src="https://www.mail-signatures.com/signature-generator/img/templates/inclusive/yt.png"
                                                                            width="22" border="0" /></a>&nbsp;</span><span><a
                                                                    href="https://www.instagram.com/techworks140/"
                                                                    rel="noopener"><img
                                                                            style="border: 0; height: 22px; width: 22px;"
                                                                            src="https://www.mail-signatures.com/signature-generator/img/templates/inclusive/it.png"
                                                                            width="22" border="0" /></a></span></p>
                                    </td>
                            </tr>
                            <tr style="width: 420px;">
                                    <td style="padding-top: 14px;" colspan="4"><a href="https://techworksworld.com/" target="_"><img
                                                            style="width: 420px; height: auto; border: 0;"
                                                            src="https://i.imgur.com/QoPxSPy.png" width="420" border="0" /></a></td>
                            </tr>
                            <tr>
                                    <td style="padding-top: 14px; text-align: justify;" colspan="4">
                                            <table style="width: 420px; font-size: 10pt; font-family: Verdana, sans-serif; background: transparent !important;"
                                                    border="0" cellspacing="0" cellpadding="0">
                                                    <tbody>
                                                            <tr>
                                                                    <td
                                                                            style="font-size: 8pt; color: #b2b2b2; line-height: 9pt; text-align: justify;">
                                                                            The content of this email is confidential and intended
                                                                            for the recipient specified in message only. It is
                                                                            strictly forbidden to share any part of this message
                                                                            with any third party,without a written consent of the
                                                                            sender. If you received this message by mistake,please
                                                                            reply to this message and follow with its deletion,so
                                                                            that we can ensure such a mistake does not occur in the
                                                                            future.</td>
                                                            </tr>
                                                    </tbody>
                                            </table>
                                    </td>
                            </tr>
                    </tbody>
            </table>`,
            attachments: [
                {
                    filename: fileName,
                    path: path.join(reportsFolderPath, fileName)
                }
            ]
        };

        const info = await transporter.sendMail(mailOptions);
        console.log('Email sent successfully:', info.messageId);
    } catch (error) {
        console.error('Error sending email:', error);
    }
}

createSeparateExcelFiles();
