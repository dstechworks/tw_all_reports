const querystring = require("querystring");
const moment = require('moment-timezone');
const nodemailer = require("nodemailer");
const dotenv = require("dotenv");
const axios = require("axios");
const path = require("path");
const xlsx = require("xlsx");
const fs = require("fs");

dotenv.config();


let timezone = 'Asia/Kolkata';
const current_date = moment().tz(timezone).format('DD-MM-YYYY');

// Email configuration
const config = {
    to: "balendu.patel@jindalsteel.com, vijay.das@jindalsteel.com",
    cc: "dhruv@techworks.co.in, sandip@techworks.co.in, rusum@techworks.co.in, hitesh.kumar@techworks.co.in",
    subject: "JSPL Raigarh Screens Notification Report - "
};

const transporter1 = nodemailer.createTransport({
    host: "smtp.dreamhost.com",
    port: 465,
    secure: true,
    auth: {
        user: "Rusum@techworks.co.in",
        pass: "E$4aFt6wEm36#AaK"
    }
});

/* -------------------- Get Access Token -------------------- */
async function getAccessToken() {
    const requestBody = {
        grant_type: "client_credentials",
        client_id: process.env.XTRAVU_SERVER_CLIENT_ID,
        client_secret: process.env.XTRAVU_SERVER_CLIENT_SECRET
    };

    const tokenEndpoint = process.env.XTRAVU_SERVER_TOKEN_END_POINT_URL;

    try {
        const response = await axios.post(
            tokenEndpoint,
            querystring.stringify(requestBody),
            {
                headers: {
                    "Content-Type": "application/x-www-form-urlencoded"
                }
            }
        );

        console.log(`New token generated for Server`);
        return response.data.access_token;

    } catch (err) {
        console.error("Error getting access token:", err.response?.data || err.message);
        throw err;
    }
}

/* -------------------- Fetch API Data -------------------- */
async function getApiData() {
    const batchSize = 10;
    let offset = 0;
    const results = [];

    const apiUrl = process.env.XTRAVU_SERVER_URL;

    let token = await getAccessToken();
    let headers = { Authorization: `Bearer ${token}` };

    try {
        while (true) {
            const response = await axios.get(
                `${apiUrl}/display?start=${offset}`,
                { headers }
            );

            const data = response.data;

            if (!Array.isArray(data) || data.length === 0) break;

            results.push(...data);
            offset += batchSize;

            console.log(`Offset :: ${offset}`);
        }

        console.log(`Total Items :: ${results.length}`);
        return results;

    } catch (error) {

        // Auto refresh token if expired
        if (error.response?.status === 401) {
            console.log(`Token expired, regenerating...`);

            token = await getAccessToken();
            headers.Authorization = `Bearer ${token}`;

            return getApiData(); // retry once
        }

        console.error(
            `Error fetching data:`,
            error.response?.data || error.message
        );
        throw error;
    }
}

/* -------------------- Generate Excel Report -------------------- */
function generateExcelReport(data) {
    // Define the folder path for saving the .xls file
    const reportsFolderPath = path.join(__dirname, 'jsplReports');

    // Check if 'jsplReports' folder exists, if not create it
    if (!fs.existsSync(reportsFolderPath)) {
        fs.mkdirSync(reportsFolderPath, { recursive: true });
    }

    // Process data to match required format
    const formattedData = data.map(item => {
        // Determine Active/Inactive based on loggedIn
        const status = (item.loggedIn === 1 || item.logged_in === 1) ? 'Active' : 'Inactive';

        // Format Date field (checking multiple possible field names)
        let dateFormatted = '';
        const dateValue = item.Date || item.date || item.createdDate || item.created_date || item.timestamp;
        if (dateValue) {
            const dateMoment = moment(dateValue);
            if (dateMoment.isValid()) {
                dateFormatted = dateMoment.tz(timezone).format('DD-MM-YYYY');
            }
        }

        return {
            'displayId': item.displayId,
            'Date': moment().tz(timezone).format('DD-MM-YYYY hh:mm A'),
            'Display': item.display,
            'Active/Inactive': status
        };
    });

    // Create a new workbook
    const wb = xlsx.utils.book_new();
    const ws = xlsx.utils.json_to_sheet(formattedData);

    // Append worksheet to workbook
    xlsx.utils.book_append_sheet(wb, ws, 'Report');

    // Define file path
    const fileName = `JSPL_Raigarh_Report_${current_date}.xls`;
    const filePath = path.join(reportsFolderPath, fileName);

    // Write the workbook to a file
    xlsx.writeFile(wb, filePath);
    console.log(`Excel file created: ${filePath}`);

    return filePath;
}

async function sendReport(filePath) {
    try {
        let currentTransporter = transporter1;
        const attachments = [];

        // Add Excel file as attachment
        if (filePath && fs.existsSync(filePath)) {
            attachments.push({
                filename: path.basename(filePath),
                path: filePath
            });
        }

        const info = await currentTransporter.sendMail({
            from: 'reports@techworks.co.in',
            cc: config.cc,
            to: config.to,
            subject: config.subject + current_date,
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
            attachments: attachments
        });

        console.log("\n");
        console.log("===============================================");
        console.log(`     MAIL SENT SUCCESSFULLY     `);
        console.log("===============================================");

        return info;
    } catch (error) {
        console.error(`Error sending email:`, error);
        throw error;
    }
}

/* -------------------- Main Runner -------------------- */
(async () => {
    try {
        console.log("\n");
        console.log("===============================================");
        console.log(`     JSPL Raigarh Report - ${current_date}     `);
        console.log("===============================================");

        // Fetch data from API
        const data = await getApiData();

        if (!data || data.length === 0) {
            console.log("No data found to generate report");
            return;
        }

        console.log(`Total Records: ${data.length}`);

        // Generate Excel report
        const filePath = generateExcelReport(data);

        // Send email with attachment
        await sendReport(filePath);

        console.log("\n Process completed successfully!");

    } catch (err) {
        console.error("Script failed:", err.message);
        console.error(err);
    }
})();
