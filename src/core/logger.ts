import winston from 'winston';
import * as path from 'path';
import * as fs from 'fs';

/**
 * Logger service for the Intune Reporting Dashboard
 * Singleton pattern ensures consistent logging across the application
 */
export class Logger {
  private static instance: Logger;
  private logger: winston.Logger;

  private constructor() {
    const logsDir = process.env.LOG_DIR || './logs';

    // Ensure logs directory exists
    if (!fs.existsSync(logsDir)) {
      fs.mkdirSync(logsDir, { recursive: true });
    }

    const logLevel = process.env.LOG_LEVEL || 'info';

    this.logger = winston.createLogger({
      level: logLevel,
      format: winston.format.combine(
        winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
        winston.format.errors({ stack: true }),
        winston.format.splat(),
        winston.format.json()
      ),
      defaultMeta: { service: 'intune-reporting' },
      transports: [
        // Write all logs to combined.log
        new winston.transports.File({
          filename: path.join(logsDir, 'combined.log'),
          maxsize: 10485760, // 10MB
          maxFiles: 5,
        }),
        // Write errors to error.log
        new winston.transports.File({
          filename: path.join(logsDir, 'error.log'),
          level: 'error',
          maxsize: 10485760,
          maxFiles: 5,
        }),
      ],
    });

    // Also log to console in development
    if (process.env.NODE_ENV !== 'production') {
      this.logger.add(
        new winston.transports.Console({
          format: winston.format.combine(
            winston.format.colorize(),
            winston.format.simple()
          ),
        })
      );
    }
  }

  public static getInstance(): Logger {
    if (!Logger.instance) {
      Logger.instance = new Logger();
    }
    return Logger.instance;
  }

  public info(message: string, meta?: any): void {
    this.logger.info(message, meta);
  }

  public error(message: string, error?: any): void {
    if (error instanceof Error) {
      this.logger.error(message, {
        error: {
          message: error.message,
          stack: error.stack,
          name: error.name,
        },
      });
    } else {
      this.logger.error(message, { error });
    }
  }

  public warn(message: string, meta?: any): void {
    this.logger.warn(message, meta);
  }

  public debug(message: string, meta?: any): void {
    this.logger.debug(message, meta);
  }

  public getLogger(): winston.Logger {
    return this.logger;
  }
}
