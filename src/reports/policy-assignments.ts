import { BaseReport } from './base-report';
import { ReportData } from '../types';
import { Logger } from '../core/logger';

const logger = Logger.getInstance();

/**
 * Policy Assignments Report
 * Shows all configuration policies and their assignments
 */
export class PolicyAssignmentsReport extends BaseReport {
  name = 'policy-assignments';
  description = 'Configuration policies and their group assignments';
  category = 'Policies';
  enabled = true;

  async execute(): Promise<ReportData> {
    logger.info('Executing Policy Assignments Report');

    try {
      // Fetch configuration policies
      const configPoliciesResponse = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceManagement/deviceConfigurations')
          .select(['id', 'displayName', 'description', 'createdDateTime', 'lastModifiedDateTime'])
          .get()
      );

      const configPolicies = await this.getAllPages(configPoliciesResponse);

      // Fetch compliance policies
      const compliancePoliciesResponse = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceManagement/deviceCompliancePolicies')
          .select(['id', 'displayName', 'description', 'createdDateTime', 'lastModifiedDateTime'])
          .get()
      );

      const compliancePolicies = await this.getAllPages(compliancePoliciesResponse);

      const data: any[] = [];

      // Process configuration policies
      for (const policy of configPolicies) {
        try {
          const assignmentsResponse = await this.retryGraphCall(() =>
            this.graphClient
              .api(`/deviceManagement/deviceConfigurations/${policy.id}/assignments`)
              .get()
          );

          const assignments = assignmentsResponse.value || [];

          data.push({
            policyName: policy.displayName || 'Unknown',
            policyType: 'Configuration',
            description: policy.description || 'No description',
            assignmentCount: assignments.length,
            createdDate: policy.createdDateTime
              ? new Date(policy.createdDateTime).toLocaleDateString()
              : 'N/A',
            lastModified: policy.lastModifiedDateTime
              ? new Date(policy.lastModifiedDateTime).toLocaleDateString()
              : 'N/A',
          });
        } catch (error) {
          logger.warn(`Failed to fetch assignments for policy ${policy.displayName}`, error);
          data.push({
            policyName: policy.displayName || 'Unknown',
            policyType: 'Configuration',
            description: policy.description || 'No description',
            assignmentCount: 'Error',
            createdDate: policy.createdDateTime
              ? new Date(policy.createdDateTime).toLocaleDateString()
              : 'N/A',
            lastModified: policy.lastModifiedDateTime
              ? new Date(policy.lastModifiedDateTime).toLocaleDateString()
              : 'N/A',
          });
        }
      }

      // Process compliance policies
      for (const policy of compliancePolicies) {
        try {
          const assignmentsResponse = await this.retryGraphCall(() =>
            this.graphClient
              .api(`/deviceManagement/deviceCompliancePolicies/${policy.id}/assignments`)
              .get()
          );

          const assignments = assignmentsResponse.value || [];

          data.push({
            policyName: policy.displayName || 'Unknown',
            policyType: 'Compliance',
            description: policy.description || 'No description',
            assignmentCount: assignments.length,
            createdDate: policy.createdDateTime
              ? new Date(policy.createdDateTime).toLocaleDateString()
              : 'N/A',
            lastModified: policy.lastModifiedDateTime
              ? new Date(policy.lastModifiedDateTime).toLocaleDateString()
              : 'N/A',
          });
        } catch (error) {
          logger.warn(`Failed to fetch assignments for policy ${policy.displayName}`, error);
          data.push({
            policyName: policy.displayName || 'Unknown',
            policyType: 'Compliance',
            description: policy.description || 'No description',
            assignmentCount: 'Error',
            createdDate: policy.createdDateTime
              ? new Date(policy.createdDateTime).toLocaleDateString()
              : 'N/A',
            lastModified: policy.lastModifiedDateTime
              ? new Date(policy.lastModifiedDateTime).toLocaleDateString()
              : 'N/A',
          });
        }
      }

      // Calculate summary statistics
      const summary = {
        totalPolicies: data.length,
        configurationPolicies: configPolicies.length,
        compliancePolicies: compliancePolicies.length,
        policiesWithAssignments: data.filter(
          (p) => typeof p.assignmentCount === 'number' && p.assignmentCount > 0
        ).length,
        policiesWithoutAssignments: data.filter(
          (p) => typeof p.assignmentCount === 'number' && p.assignmentCount === 0
        ).length,
      };

      logger.info('Policy Assignments Report completed', { policyCount: data.length });

      return {
        metadata: this.createMetadata(this.name, data.length),
        data,
        summary,
      };
    } catch (error) {
      logger.error('Failed to execute Policy Assignments Report', error);
      throw error;
    }
  }
}
