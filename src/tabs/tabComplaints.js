const { google } = require('googleapis');
const moment = require('moment-timezone');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const nodemailer = require("nodemailer");

// Function to get the dynamic path to credentials.json
function getCredentialsPath() {
    // Get the directory of the current script
    const currentScriptDir = __dirname;
    // Navigate to the project root (2 levels up from current script)
    const projectRoot = path.resolve(currentScriptDir, '../..');
    // Return the path to credentials.json in the project root
    return path.join(projectRoot, 'credentials.json');
}
