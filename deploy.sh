#!/bin/bash

echo "===== DEPLOY STARTED at $(date) =====" >> /tmp/deploy.log

cd ~/tw_all_reports || {
  echo "❌ ERROR: Could not cd into project folder" >> /tmp/deploy.log
  exit 1
}

# Load NVM (for node/npm access)
export NVM_DIR="$HOME/.nvm"
[ -s "$NVM_DIR/nvm.sh" ] && \. "$NVM_DIR/nvm.sh"

# Pull latest code
echo "Pulling code from origin/production..." >> /tmp/deploy.log
git pull origin production >> /tmp/deploy.log 2>&1

# Remove and reinstall dependencies
echo "Removing node_modules..." >> /tmp/deploy.log
rm -rf node_modules >> /tmp/deploy.log 2>&1

echo "Installing dependencies..." >> /tmp/deploy.log
npm install >> /tmp/deploy.log 2>&1

# Restart the app with PM2 (process ID 7)
echo "Restarting app with PM2 (ID 7)..." >> /tmp/deploy.log
pm2 stop 7 >> /tmp/deploy.log 2>&1
pm2 restart 7 >> /tmp/deploy.log 2>&1
pm2 save >> /tmp/deploy.log 2>&1

echo "===== DEPLOY COMPLETED at $(date) =====" >> /tmp/deploy.log