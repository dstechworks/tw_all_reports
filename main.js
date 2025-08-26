
const { exec } = require('child_process');
const moment = require('moment-timezone');
const cron = require('node-cron');

const TIMEZONE = 'Asia/Kolkata';

// Print message with current time
function logWithTime(message) {
    const time = moment().tz(TIMEZONE).format('YYYY-MM-DD HH:mm:ss');
    console.log(`[${time}] ${message}`);
}

// Run a script using node
function runScript(name, path) {
    logWithTime(`Starting ${name}...`);
    exec(`node ${path}`, (error, stdout, stderr) => {
        if (error) {
            logWithTime(`${name} ERROR: ${error.message}`);
            return;
        }
        if (stderr) {
            logWithTime(`${name} STDERR: ${stderr}`);
        }
        logWithTime(`${name} OUTPUT: ${stdout.trim()}`);
    });
}

// Every Monday @ 11:58 AM: Run MPDU Complaints
cron.schedule('58 11 * * 1', () => {
    runScript('MPDU Complaints', 'src/mpdu/mpdu_43vertical_complaint_report/mpduComplaints.js'); // Monday 11:58 AM
}, { timezone: TIMEZONE });

// 15th day of month and last day of month @ 12:00 PM: Run Tab Complaints
cron.schedule('0 12 15 * *', () => {
    runScript('Tab Complaints', 'src/tabs/tabComplaints.js'); // 15th day of month at 12 PM
}, { timezone: TIMEZONE });

// Last day of month @ 12:00 PM: Run Tab Complaints
cron.schedule('0 12 28-31 * *', () => {
    const now = moment().tz(TIMEZONE);
    const lastDayOfMonth = now.endOf('month').date();
    if (now.date() === lastDayOfMonth) {
        runScript('Tab Complaints', 'src/tabs/tabComplaints.js'); // Last day of month at 12 PM
    }
}, { timezone: TIMEZONE });

// Run a script manually if i want to runscript manually without using cron
runScript('MPDU Complaints', 'src/mpdu/mpdu_43vertical_complaint_report/mpduComplaints.js');
// runScript('Tab Complaints', 'src/tabs/tabComplaints.js');
// runScript('MPDU 43 Vertical', 'src/mpdu/mpdu_43vertical_month_end_report/mpdu.js');
// runScript('MPDU 43 Vertical', 'src/mpdu/mpdu_43vertical_month_end_report/43vertical.js');