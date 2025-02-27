import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

class Logger {
    constructor() {
        // Create logs directory if it doesn't exist
        const logsDir = path.join(__dirname, 'logs');
        if (!fs.existsSync(logsDir)) {
            fs.mkdirSync(logsDir);
        }

        // Initialize log files with timestamps
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        this.successLogPath = path.join(logsDir, `success_${timestamp}.log`);
        this.errorLogPath = path.join(logsDir, `error_${timestamp}.log`);
        this.combinedLogPath = path.join(logsDir, `combined_${timestamp}.log`);
    }

    formatMessage(level, message, details = null) {
        const timestamp = new Date().toISOString();
        let logMessage = `[${timestamp}] [${level}] ${message}`;
        if (details) {
            logMessage += `\nDetails: ${JSON.stringify(details, null, 2)}\n`;
        }
        return logMessage + '\n';
    }

    success(message, details = null) {
        const logMessage = this.formatMessage('SUCCESS', message, details);
        fs.appendFileSync(this.successLogPath, logMessage);
        fs.appendFileSync(this.combinedLogPath, logMessage);
        console.log('\x1b[32m%s\x1b[0m', message); // Green color
    }

    error(message, error = null) {
        const details = error ? {
            message: error.message,
            stack: error.stack,
            ...(error.details || {})
        } : null;
        const logMessage = this.formatMessage('ERROR', message, details);
        fs.appendFileSync(this.errorLogPath, logMessage);
        fs.appendFileSync(this.combinedLogPath, logMessage);
        console.error('\x1b[31m%s\x1b[0m', message); // Red color
    }

    info(message, details = null) {
        const logMessage = this.formatMessage('INFO', message, details);
        console.log('Writing log:', logMessage); // Debug line
        fs.appendFileSync(this.combinedLogPath, logMessage);
        fs.appendFileSync(this.successLogPath, logMessage);
        console.log('\x1b[36m%s\x1b[0m', message); // Cyan color
    }
}

export default new Logger(); 