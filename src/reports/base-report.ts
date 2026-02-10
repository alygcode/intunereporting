import { Client } from '@microsoft/microsoft-graph-client';
import { Report, ReportData, AppConfig } from '../types';
import { Logger } from '../core/logger';

const logger = Logger.getInstance();

/**
 * Base Report class - provides common functionality for all reports
 */
export abstract class BaseReport implements Report {
  protected graphClient: Client;
  protected config: AppConfig;

  abstract name: string;
  abstract description: string;
  abstract category: string;
  abstract enabled: boolean;

  constructor(graphClient: Client, config: AppConfig) {
    this.graphClient = graphClient;
    this.config = config;
  }

  /**
   * Execute the report - must be implemented by subclasses
   */
  abstract execute(): Promise<ReportData>;

  /**
   * Create report metadata
   */
  protected createMetadata(reportName: string, recordCount: number, parameters?: Record<string, any>): ReportData['metadata'] {
    return {
      reportName,
      generatedAt: new Date().toISOString(),
      generatedBy: 'Intune Reporting Dashboard',
      recordCount,
      parameters,
    };
  }

  /**
   * Handle Graph API pagination
   */
  protected async getAllPages<T>(initialResponse: any): Promise<T[]> {
    let results: T[] = initialResponse.value || [];
    let nextLink = initialResponse['@odata.nextLink'];

    while (nextLink) {
      try {
        const response = await this.graphClient.api(nextLink).get();
        results = results.concat(response.value || []);
        nextLink = response['@odata.nextLink'];
      } catch (error) {
        logger.error('Error fetching next page', error);
        break;
      }
    }

    return results;
  }

  /**
   * Retry a Graph API call with exponential backoff
   */
  protected async retryGraphCall<T>(
    apiCall: () => Promise<T>,
    maxRetries: number = 3,
    baseDelay: number = 1000
  ): Promise<T> {
    for (let attempt = 0; attempt < maxRetries; attempt++) {
      try {
        return await apiCall();
      } catch (error: any) {
        if (attempt === maxRetries - 1) {
          throw error;
        }

        const statusCode = error?.statusCode;
        if (statusCode === 429 || statusCode >= 500) {
          const delay = baseDelay * Math.pow(2, attempt);
          logger.warn(`Rate limited or server error, retrying in ${delay}ms`, { attempt, statusCode });
          await this.sleep(delay);
        } else {
          throw error;
        }
      }
    }

    throw new Error('Max retries exceeded');
  }

  /**
   * Sleep helper
   */
  protected sleep(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }
}
