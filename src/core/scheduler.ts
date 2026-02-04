import { CronJob } from 'cron';
import { ReportOrchestrator } from './orchestrator';
import { AppConfig, ScheduleDefinition } from '../types';
import { Logger } from './logger';

const logger = Logger.getInstance();

/**
 * Scheduler Manager for automated report generation
 * Handles cron-based scheduling of reports
 */
export class SchedulerManager {
  private orchestrator: ReportOrchestrator;
  private config: AppConfig;
  private jobs: Map<string, CronJob>;

  constructor(orchestrator: ReportOrchestrator, config: AppConfig) {
    this.orchestrator = orchestrator;
    this.config = config;
    this.jobs = new Map();
  }

  /**
   * Start the scheduler and all enabled jobs
   */
  start(): void {
    if (!this.config.scheduler?.enabled) {
      logger.warn('Scheduler is disabled in configuration');
      return;
    }

    logger.info('Starting scheduler');

    const { schedules, timezone } = this.config.scheduler;

    schedules.forEach((schedule) => {
      if (!schedule.enabled) {
        logger.info(`Skipping disabled schedule: ${schedule.name}`);
        return;
      }

      try {
        const job = this.createJob(schedule, timezone);
        this.jobs.set(schedule.name, job);
        job.start();

        logger.info(`Scheduled job started: ${schedule.name}`, {
          cron: schedule.cron,
          reports: schedule.reports,
          timezone,
        });
      } catch (error) {
        logger.error(`Failed to create scheduled job: ${schedule.name}`, error);
      }
    });

    logger.info(`Started ${this.jobs.size} scheduled jobs`);
  }

  /**
   * Stop all scheduled jobs
   */
  stop(): void {
    logger.info('Stopping scheduler');

    this.jobs.forEach((job, name) => {
      job.stop();
      logger.info(`Stopped scheduled job: ${name}`);
    });

    this.jobs.clear();
    logger.info('Scheduler stopped');
  }

  /**
   * Create a cron job for a schedule definition
   */
  private createJob(schedule: ScheduleDefinition, timezone: string): CronJob {
    return new CronJob(
      schedule.cron,
      async () => {
        logger.info(`Executing scheduled job: ${schedule.name}`);
        const startTime = Date.now();

        try {
          const results = await this.orchestrator.runReports(
            schedule.reports,
            schedule.formats,
            this.config.output.directory
          );

          const duration = Date.now() - startTime;
          const successful = results.filter((r) => r.success).length;
          const failed = results.filter((r) => !r.success).length;

          logger.info(`Scheduled job completed: ${schedule.name}`, {
            duration,
            successful,
            failed,
            total: results.length,
          });

          // Log individual report results
          results.forEach((result) => {
            if (result.success) {
              logger.info(`Report generated: ${result.reportName}`, {
                format: result.format,
                outputPath: result.outputPath,
                recordCount: result.recordCount,
              });
            } else {
              logger.error(`Report failed: ${result.reportName}`, {
                error: result.error,
              });
            }
          });
        } catch (error) {
          logger.error(`Scheduled job failed: ${schedule.name}`, error);
        }
      },
      null, // onComplete
      false, // start
      timezone
    );
  }

  /**
   * Get status of all scheduled jobs
   */
  getStatus(): Array<{ name: string; running: boolean; nextRun?: Date }> {
    const status: Array<{ name: string; running: boolean; nextRun?: Date }> = [];

    this.jobs.forEach((job, name) => {
      status.push({
        name,
        running: job.running,
        nextRun: job.nextDate()?.toJSDate(),
      });
    });

    return status;
  }

  /**
   * Manually trigger a scheduled job
   */
  async triggerJob(jobName: string): Promise<void> {
    const job = this.jobs.get(jobName);
    if (!job) {
      throw new Error(`Job not found: ${jobName}`);
    }

    logger.info(`Manually triggering job: ${jobName}`);
    await job.fireOnTick();
  }
}
