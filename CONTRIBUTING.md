# Contributing to Intune Reporting Dashboard

Thank you for your interest in contributing to the Intune Reporting Dashboard! This document provides guidelines and instructions for contributing.

## Getting Started

1. Fork the repository
2. Clone your fork locally
3. Install dependencies: `npm install`
4. Build the project: `npm run build`
5. Create a feature branch: `git checkout -b feature/your-feature-name`

## Development Setup

### Prerequisites
- Node.js 16.0.0 or higher
- TypeScript 5.x
- Access to a test Azure AD tenant with Intune

### Environment Setup
1. Copy `.env.example` to `.env`
2. Fill in your Azure AD test credentials
3. Run `npm run dev -- init` to create configuration file
4. Update configuration with your test settings

## Code Style

- Use TypeScript strict mode
- Follow existing code formatting
- Use meaningful variable and function names
- Add comments for complex logic
- Keep functions small and focused

## Adding a New Report

To add a new report to the dashboard:

1. Create a new file in `src/reports/` directory:

```typescript
import { BaseReport } from './base-report';
import { ReportData } from '../types';
import { Logger } from '../core/logger';

const logger = Logger.getInstance();

export class YourNewReport extends BaseReport {
  name = 'your-report-name';
  description = 'Description of your report';
  category = 'Category';
  enabled = true;

  async execute(): Promise<ReportData> {
    logger.info('Executing Your New Report');

    try {
      // Fetch data from Graph API
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api('/your/graph/endpoint')
          .select(['field1', 'field2'])
          .get()
      );

      const items = await this.getAllPages(response);

      // Transform data
      const data = items.map((item: any) => ({
        field1: item.field1,
        field2: item.field2,
      }));

      // Calculate summary
      const summary = {
        total: data.length,
      };

      return {
        metadata: this.createMetadata(this.name, data.length),
        data,
        summary,
      };
    } catch (error) {
      logger.error('Failed to execute Your New Report', error);
      throw error;
    }
  }
}
```

2. Register the report in `src/core/orchestrator.ts`:

```typescript
import { YourNewReport } from '../reports/your-new-report';

// In registerReports() method:
const reportClasses = [
  // ... existing reports
  YourNewReport,
];
```

3. Add tests for your report
4. Update documentation

## Testing

### Running Tests
```bash
npm test
```

### Writing Tests
- Add tests for new functionality
- Ensure existing tests pass
- Aim for good code coverage
- Mock external dependencies (Graph API calls)

## Documentation

- Update README.md for user-facing changes
- Add JSDoc comments to public APIs
- Update CHANGELOG.md
- Include examples for new features

## Pull Request Process

1. Ensure all tests pass
2. Update documentation
3. Add yourself to contributors if first-time contributor
4. Create a pull request with clear description:
   - What changes were made
   - Why the changes were needed
   - How to test the changes
   - Screenshots (if applicable)

### Pull Request Template

```markdown
## Description
Brief description of changes

## Type of Change
- [ ] Bug fix
- [ ] New feature
- [ ] Breaking change
- [ ] Documentation update

## Testing
How has this been tested?

## Checklist
- [ ] Code follows project style guidelines
- [ ] Self-review completed
- [ ] Comments added for complex code
- [ ] Documentation updated
- [ ] Tests added/updated
- [ ] All tests passing
```

## Code Review

- Be respectful and constructive
- Focus on the code, not the person
- Explain your reasoning
- Be open to feedback

## Reporting Bugs

When reporting bugs, include:
- Clear description of the issue
- Steps to reproduce
- Expected behavior
- Actual behavior
- Environment details (Node version, OS, etc.)
- Relevant logs or error messages

## Suggesting Enhancements

For feature requests:
- Explain the use case
- Describe the proposed solution
- Consider alternatives
- Assess impact on existing functionality

## License

By contributing, you agree that your contributions will be licensed under the MIT License.

## Questions?

Feel free to create an issue for any questions or clarifications needed.

Thank you for contributing!
