import * as fs from 'fs/promises';
import * as path from 'path';
import { Parser } from 'json2csv';
import { ReportData } from '../types';
import { Logger } from '../core/logger';

const logger = Logger.getInstance();

/**
 * Output Formatter - Handles formatting report data to different output formats
 */
export class OutputFormatter {
  /**
   * Format and save report data in the specified format
   */
  static async format(
    reportData: ReportData,
    format: 'json' | 'csv' | 'html',
    outputDir: string,
    includeTimestamp: boolean = true
  ): Promise<string> {
    const timestamp = includeTimestamp ? `-${this.getTimestamp()}` : '';
    const fileName = `${this.sanitizeFileName(reportData.metadata.reportName)}${timestamp}`;
    const outputPath = path.join(outputDir, `${fileName}.${format}`);

    switch (format) {
      case 'json':
        await this.formatJson(reportData, outputPath);
        break;
      case 'csv':
        await this.formatCsv(reportData, outputPath);
        break;
      case 'html':
        await this.formatHtml(reportData, outputPath);
        break;
      default:
        throw new Error(`Unsupported output format: ${format}`);
    }

    return outputPath;
  }

  /**
   * Format data as JSON
   */
  private static async formatJson(reportData: ReportData, outputPath: string): Promise<void> {
    const jsonContent = JSON.stringify(reportData, null, 2);
    await fs.writeFile(outputPath, jsonContent, 'utf-8');
    logger.debug(`JSON report saved to: ${outputPath}`);
  }

  /**
   * Format data as CSV
   */
  private static async formatCsv(reportData: ReportData, outputPath: string): Promise<void> {
    if (!reportData.data || reportData.data.length === 0) {
      // Write empty CSV with headers from metadata
      await fs.writeFile(outputPath, 'No data available\n', 'utf-8');
      logger.debug(`Empty CSV report saved to: ${outputPath}`);
      return;
    }

    try {
      const parser = new Parser({
        flatten: true,
        unwind: [],
      });
      const csv = parser.parse(reportData.data);

      // Add metadata as comments at the top
      const metadata = [
        `# Report: ${reportData.metadata.reportName}`,
        `# Generated: ${reportData.metadata.generatedAt}`,
        `# Records: ${reportData.metadata.recordCount}`,
        '',
      ].join('\n');

      await fs.writeFile(outputPath, metadata + csv, 'utf-8');
      logger.debug(`CSV report saved to: ${outputPath}`);
    } catch (error) {
      logger.error('Failed to generate CSV', error);
      throw new Error(`CSV generation failed: ${(error as Error).message}`);
    }
  }

  /**
   * Format data as HTML
   */
  private static async formatHtml(reportData: ReportData, outputPath: string): Promise<void> {
    const html = this.generateHtmlReport(reportData);
    await fs.writeFile(outputPath, html, 'utf-8');
    logger.debug(`HTML report saved to: ${outputPath}`);
  }

  /**
   * Generate HTML report with styling
   */
  private static generateHtmlReport(reportData: ReportData): string {
    const { metadata, data, summary } = reportData;

    const htmlHeader = `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${metadata.reportName} - Intune Report</title>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        h1 {
            color: #0078d4;
            margin-top: 0;
            border-bottom: 3px solid #0078d4;
            padding-bottom: 10px;
        }
        .metadata {
            background-color: #f0f0f0;
            padding: 15px;
            border-radius: 4px;
            margin-bottom: 20px;
        }
        .metadata p {
            margin: 5px 0;
        }
        .summary {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin-bottom: 20px;
        }
        .summary-card {
            background-color: #e3f2fd;
            padding: 15px;
            border-radius: 4px;
            border-left: 4px solid #0078d4;
        }
        .summary-card h3 {
            margin: 0 0 5px 0;
            font-size: 14px;
            color: #666;
        }
        .summary-card p {
            margin: 0;
            font-size: 24px;
            font-weight: bold;
            color: #0078d4;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        th {
            background-color: #0078d4;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: 600;
        }
        td {
            padding: 10px 12px;
            border-bottom: 1px solid #ddd;
        }
        tr:hover {
            background-color: #f5f5f5;
        }
        .footer {
            margin-top: 30px;
            padding-top: 15px;
            border-top: 1px solid #ddd;
            color: #666;
            font-size: 12px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>${metadata.reportName}</h1>

        <div class="metadata">
            <p><strong>Generated:</strong> ${metadata.generatedAt}</p>
            <p><strong>Generated By:</strong> ${metadata.generatedBy}</p>
            <p><strong>Total Records:</strong> ${metadata.recordCount}</p>
        </div>
`;

    let summaryHtml = '';
    if (summary && Object.keys(summary).length > 0) {
      summaryHtml = '<div class="summary">';
      Object.entries(summary).forEach(([key, value]) => {
        summaryHtml += `
            <div class="summary-card">
                <h3>${this.formatLabel(key)}</h3>
                <p>${value}</p>
            </div>
        `;
      });
      summaryHtml += '</div>';
    }

    let tableHtml = '';
    if (data && data.length > 0) {
      const headers = Object.keys(data[0]);
      tableHtml = '<table><thead><tr>';
      headers.forEach((header) => {
        tableHtml += `<th>${this.formatLabel(header)}</th>`;
      });
      tableHtml += '</tr></thead><tbody>';

      data.forEach((row) => {
        tableHtml += '<tr>';
        headers.forEach((header) => {
          const value = row[header];
          tableHtml += `<td>${this.formatValue(value)}</td>`;
        });
        tableHtml += '</tr>';
      });

      tableHtml += '</tbody></table>';
    } else {
      tableHtml = '<p>No data available for this report.</p>';
    }

    const htmlFooter = `
        <div class="footer">
            <p>Intune Reporting Dashboard | Report generated on ${new Date().toLocaleString()}</p>
        </div>
    </div>
</body>
</html>
`;

    return htmlHeader + summaryHtml + tableHtml + htmlFooter;
  }

  /**
   * Format label for display
   */
  private static formatLabel(label: string): string {
    return label
      .replace(/([A-Z])/g, ' $1')
      .replace(/^./, (str) => str.toUpperCase())
      .trim();
  }

  /**
   * Format value for HTML display
   */
  private static formatValue(value: any): string {
    if (value === null || value === undefined) {
      return '<em>N/A</em>';
    }
    if (typeof value === 'boolean') {
      return value ? 'Yes' : 'No';
    }
    if (typeof value === 'object') {
      return JSON.stringify(value);
    }
    return String(value);
  }

  /**
   * Get timestamp string for file naming
   */
  private static getTimestamp(): string {
    const now = new Date();
    return now
      .toISOString()
      .replace(/[:.]/g, '-')
      .replace('T', '_')
      .split('.')[0];
  }

  /**
   * Sanitize file name
   */
  private static sanitizeFileName(name: string): string {
    return name.toLowerCase().replace(/[^a-z0-9-]/g, '-');
  }
}
