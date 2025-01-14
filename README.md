# Microsoft Graph Pipedream Components

## Overview

This package contains Pipedream components for interacting with the [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/overview), focusing on subscription management and real-time notifications.

## Key Features

- **Subscription Management**: Create, monitor, and manage Microsoft Graph subscriptions
- **Real-time Notifications**: Handle webhook notifications for resource changes
- **Automatic Renewal**: Keep subscriptions active with automatic renewal
- **Health Monitoring**: Track subscription health and handle errors gracefully

## Components

### Sources

1. **Subscription Creator** (`microsoft_graph-subscription-creator`)
   - Creates and initializes Microsoft Graph subscriptions
   - Supports multiple resource types (messages, calendar, contacts)
   - Includes validation and error handling

2. **Subscription Monitor** (`microsoft_graph-subscription-monitor`)
   - Monitors subscription health and expiration
   - Automatically renews subscriptions before expiry
   - Handles cleanup of old data

3. **Notification Handler** (`microsoft_graph-notification-handler`)
   - Processes incoming webhook notifications
   - Validates notification authenticity
   - Handles lifecycle events and resource changes

## Getting Started

1. **Authentication**
   - Create a Microsoft Azure AD application
   - Configure required permissions (Mail.Read, etc.)
   - Generate client credentials

2. **Configuration**
   - Set up webhook endpoint (HTTPS required)
   - Configure subscription parameters
   - Set up monitoring schedule

3. **Usage**
   ```javascript
   // Example: Creating a subscription
   const subscription = await this.microsoft_graph.createSubscription($, {
     resource: "me/mailFolders('Inbox')/messages",
     webhookUrl: "https://your-webhook.com/endpoint",
     expirationDateTime: "2024-01-20T11:00:00.000Z",
     includeResourceData: false
   });
   ```

## Common Issues and Solutions

1. **Webhook Validation**
   - Ensure webhook URL is HTTPS
   - Implement proper validation response
   - Handle lifecycle notifications

2. **Subscription Expiration**
   - Monitor subscription status
   - Renew before expiration (recommended: 12 hours before)
   - Handle reauthorization requirements

3. **Rate Limiting**
   - Implement proper backoff strategy
   - Monitor API usage
   - Handle throttling responses

## Contributing

Contributions are welcome! Please read our [contributing guidelines](https://pipedream.com/docs/components/guidelines/) before submitting changes.

## Support

- [Pipedream Documentation](https://pipedream.com/docs)
- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/overview)
- [Pipedream Community](https://pipedream.com/community)

## License

MIT © [Pipedream, Inc.](https://pipedream.com) 