import * as fs from 'fs/promises';
import * as path from 'path';
import YAML from 'yaml';
import { AppConfig, ValidationResult } from '../types';
import { Logger } from './logger';

const logger = Logger.getInstance();

/**
 * Configuration Manager for the Intune Reporting Dashboard
 * Handles loading, validation, and creation of configuration files
 */
export class ConfigManager {
  /**
   * Load configuration from a YAML file
   */
  static async load(configPath: string): Promise<AppConfig> {
    try {
      logger.info(`Loading configuration from: ${configPath}`);

      const content = await fs.readFile(configPath, 'utf-8');
      const config = YAML.parse(content) as AppConfig;

      // Merge with environment variables
      const mergedConfig = this.mergeWithEnv(config);

      logger.info('Configuration loaded successfully');
      return mergedConfig;
    } catch (error) {
      logger.error('Failed to load configuration', error);
      throw new Error(`Failed to load configuration: ${(error as Error).message}`);
    }
  }

  /**
   * Merge configuration with environment variables
   */
  private static mergeWithEnv(config: AppConfig): AppConfig {
    return {
      ...config,
      authentication: {
        ...config.authentication,
        tenantId: process.env.AZURE_TENANT_ID || config.authentication.tenantId,
        clientId: process.env.AZURE_CLIENT_ID || config.authentication.clientId,
        clientSecret: process.env.AZURE_CLIENT_SECRET || config.authentication.clientSecret,
      },
    };
  }

  /**
   * Validate configuration structure and required fields
   */
  static validate(config: AppConfig): ValidationResult {
    const errors: string[] = [];

    // Validate authentication
    if (!config.authentication) {
      errors.push('Authentication configuration is missing');
    } else {
      if (!config.authentication.tenantId) {
        errors.push('Tenant ID is required');
      }
      if (!config.authentication.clientId) {
        errors.push('Client ID is required');
      }
      if (config.authentication.authMethod === 'clientSecret' && !config.authentication.clientSecret) {
        errors.push('Client secret is required for clientSecret auth method');
      }
      if (config.authentication.authMethod === 'certificate') {
        if (!config.authentication.certificatePath || !config.authentication.certificateThumbprint) {
          errors.push('Certificate path and thumbprint are required for certificate auth method');
        }
      }
    }

    // Validate output configuration
    if (!config.output) {
      errors.push('Output configuration is missing');
    } else {
      if (!config.output.directory) {
        errors.push('Output directory is required');
      }
      const validFormats = ['json', 'csv', 'html'];
      if (!validFormats.includes(config.output.defaultFormat)) {
        errors.push(`Invalid default format: ${config.output.defaultFormat}`);
      }
    }

    // Validate scheduler configuration
    if (config.scheduler?.enabled) {
      if (!config.scheduler.schedules || config.scheduler.schedules.length === 0) {
        errors.push('At least one schedule is required when scheduler is enabled');
      } else {
        config.scheduler.schedules.forEach((schedule, index) => {
          if (!schedule.name) {
            errors.push(`Schedule ${index + 1} is missing a name`);
          }
          if (!schedule.cron) {
            errors.push(`Schedule "${schedule.name || index + 1}" is missing cron expression`);
          }
          if (!schedule.reports || schedule.reports.length === 0) {
            errors.push(`Schedule "${schedule.name || index + 1}" has no reports configured`);
          }
        });
      }
    }

    return {
      valid: errors.length === 0,
      errors,
    };
  }

  /**
   * Create a default configuration file
   */
  static async createDefaultConfig(outputPath: string): Promise<void> {
    const defaultConfig: AppConfig = {
      authentication: {
        tenantId: 'YOUR_TENANT_ID',
        clientId: 'YOUR_CLIENT_ID',
        clientSecret: 'YOUR_CLIENT_SECRET',
        authMethod: 'clientSecret',
      },
      reports: {
        enabled: [
          'device-inventory',
          'compliance-summary',
          'application-inventory',
          'user-devices',
          'policy-assignments',
        ],
        disabled: [],
        settings: {
          includeInactiveDevices: false,
          daysInactive: 30,
        },
      },
      output: {
        defaultFormat: 'json',
        directory: './reports',
        includeTimestamp: true,
        compression: false,
      },
      scheduler: {
        enabled: true,
        timezone: 'UTC',
        schedules: [
          {
            name: 'Daily Reports',
            cron: '0 6 * * *',
            reports: ['device-inventory', 'compliance-summary'],
            formats: ['json', 'csv'],
            enabled: true,
          },
          {
            name: 'Weekly Full Report',
            cron: '0 8 * * 1',
            reports: ['all'],
            formats: ['json', 'csv', 'html'],
            enabled: true,
          },
        ],
      },
      logging: {
        level: 'info',
        file: './logs/intune-reports.log',
        console: true,
        maxSize: '10m',
        maxFiles: 5,
      },
    };

    const yamlContent = YAML.stringify(defaultConfig);
    await fs.writeFile(outputPath, yamlContent, 'utf-8');
    logger.info(`Default configuration created at: ${outputPath}`);
  }
}
