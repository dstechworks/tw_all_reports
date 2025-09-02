const path = require('path');

/**
 * Get the dynamic path to credentials.json from any script location
 * This function automatically determines the correct path regardless of script location
 * @returns {string} - Absolute path to credentials.json
 */
function getCredentialsPath() {
    // Get the directory of the current script
    const currentScriptDir = __dirname;
    
    // Navigate to the project root from the utility directory (3 levels up)
    const projectRoot = path.resolve(currentScriptDir, '../../..');
    
    // Return the path to credentials.json in the project root
    return path.join(projectRoot, 'credentials.json');
}

module.exports = {
    getCredentialsPath
};
