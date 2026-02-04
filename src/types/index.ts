/**
 * Type definitions for the Intune Reporting Dashboard
 */

export interface AppConfig {
  authentication: AuthConfig;
  reports: ReportConfig;
  output: OutputConfig;
  scheduler: SchedulerConfig;
  logging: LoggingConfig;
}

export interface AuthConfig {
  tenantId: string;
  clientId: string;
  clientSecret?: string;
  certificatePath?: string;
  certificateThumbprint?: string;
  authMethod: 'clientSecret' | 'certificate' | 'interactive';
}

export interface ReportConfig {
  enabled: string[];
  disabled: string[];
  settings: Record<string, any>;
}

export interface OutputConfig {
  defaultFormat: 'json' | 'csv' | 'html';
  directory: string;
  includeTimestamp: boolean;
  compression: boolean;
}

export interface SchedulerConfig {
  enabled: boolean;
  timezone: string;
  schedules: ScheduleDefinition[];
}

export interface ScheduleDefinition {
  name: string;
  cron: string;
  reports: string[];
  formats: string[];
  enabled: boolean;
}

export interface LoggingConfig {
  level: 'error' | 'warn' | 'info' | 'debug';
  file: string;
  console: boolean;
  maxSize: string;
  maxFiles: number;
}

export interface Report {
  name: string;
  description: string;
  category: string;
  enabled: boolean;
  execute: () => Promise<ReportData>;
}

export interface ReportData {
  metadata: ReportMetadata;
  data: any[];
  summary?: Record<string, any>;
}

export interface ReportMetadata {
  reportName: string;
  generatedAt: string;
  generatedBy: string;
  recordCount: number;
  parameters?: Record<string, any>;
}

export interface ReportResult {
  reportName: string;
  success: boolean;
  format: string;
  outputPath?: string;
  error?: string;
  duration: number;
  recordCount?: number;
}

export interface SystemStatus {
  authenticated: boolean;
  tenantId?: string;
  reportsAvailable: number;
  reportsEnabled: number;
  schedulerEnabled: boolean;
  defaultFormat: string;
  outputDirectory: string;
  lastRun?: string;
}

export interface ValidationResult {
  valid: boolean;
  errors: string[];
  warnings?: string[];
}
