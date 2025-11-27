import { createLogger, format, transports } from "winston";

const { combine, timestamp, printf, colorize } = format;

// Custom log format
const logFormat = printf(({ level, message, timestamp }) => {
  return `${timestamp} [${level}]: ${message}`;
});

// Winston logger instance
const logger = createLogger({
  level: "info",
  format: combine(
    colorize(),
    timestamp({ format: "YYYY-MM-DD HH:mm:ss" }),           
    logFormat            
  ),
  transports: [
    new transports.Console(),
    // new transports.File({ filename: "logs/error.log", level: "error" }),             
  ],
});

export default logger;
