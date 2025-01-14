import { SUBSCRIPTION_PROPS, SUBSCRIPTION_METHODS } from "../../common/subscription.mjs";

/**
 * Source component that handles Microsoft Graph change notifications
 */
export default {
  key: "microsoft_graph-notification-handler",
  name: "Microsoft Graph Notification Handler",
  description: "Processes change notifications from Microsoft Graph subscriptions",
  version: "0.0.1",
  type: "source",
  
  props: {
    microsoft_graph: {
      type: "app",
      app: "microsoft_graph"
    },
    data_store: {
      type: "data_store",
      label: "Data Store",
      description: "Store for managing subscription state and metadata"
    },
    ...SUBSCRIPTION_PROPS,
    http: {
      type: "$.interface.http",
      customResponse: true
    },
    maxRetries: {
      type: "integer",
      label: "Max Retries",
      description: "Maximum number of retries for failed API calls",
      optional: true,
      default: 3,
      min: 1,
      max: 10
    },
    rateLimitPerSecond: {
      type: "integer",
      label: "Rate Limit",
      description: "Number of API calls allowed per second",
      optional: true,
      default: 10,
      min: 1,
      max: 30
    }
  },

  methods: {
    ...SUBSCRIPTION_METHODS,

    /**
     * Validates notification signature
     */
    async validateNotification(body, headers) {
      // Implement validation logic based on Microsoft Graph requirements
      if (!body || !body.value || !Array.isArray(body.value)) {
        throw new Error("Invalid notification format");
      }

      const subscription = await this.getStoredSubscription();
      if (!subscription) {
        throw new Error("No subscription found for validation");
      }

      // Validate clientState if present
      if (body.clientState && body.clientState !== subscription.clientState) {
        throw new Error("Invalid client state");
      }

      return true;
    },

    /**
     * Processes a change notification
     */
    async processChange(change, subscription, $) {
      // Acquire lock for the specific resource
      const hasLock = await this.acquireLock(change.resourceData?.id || change.resourceId);
      if (!hasLock) {
        console.log("Another instance is processing this change");
        return;
      }

      try {
        let resourceData = change.resourceData;
        
        // Fetch full resource data if not included in notification
        if (!resourceData && change.resourceId) {
          resourceData = await this.microsoft_graph.getMessage($, change.resourceId);
        }

        // Process the change based on type
        const processedChange = {
          id: change.resourceId,
          changeType: change.changeType,
          resource: subscription.resource,
          data: resourceData,
          subscriptionId: subscription.id,
          tenantId: change.tenantId,
          timestamp: new Date().toISOString()
        };

        // Store processed change
        await this.data_store.set(`change:${change.resourceId}`, {
          ...processedChange,
          processed: new Date().toISOString()
        });

        return processedChange;
      } finally {
        await this.releaseLock(change.resourceData?.id || change.resourceId);
      }
    },

    /**
     * Handles subscription lifecycle events
     */
    async handleLifecycleEvent(notification, subscription, $) {
      switch (notification.lifecycleEvent) {
        case "reauthorizationRequired":
          await this.reportHealth("warning", {
            action: "reauthorization_required",
            subscriptionId: subscription.id
          });
          break;

        case "subscriptionRemoved":
          await this.reportHealth("error", {
            action: "subscription_removed",
            subscriptionId: subscription.id,
            reason: notification.reason
          });
          break;

        case "missed":
          await this.reportHealth("warning", {
            action: "notifications_missed",
            subscriptionId: subscription.id,
            reason: notification.reason
          });
          break;
      }

      return {
        event: notification.lifecycleEvent,
        subscriptionId: subscription.id,
        timestamp: new Date().toISOString()
      };
    }
  },

  async run({ steps, $ }) {
    const { body, headers } = steps.trigger.event;

    try {
      // Validate the notification
      await this.validateNotification(body, headers);

      // Get current subscription
      const subscription = await this.getStoredSubscription();

      // Process notifications
      const results = await Promise.all(
        body.value.map(async (notification) => {
          try {
            // Handle lifecycle events
            if (notification.lifecycleEvent) {
              return await this.handleLifecycleEvent(notification, subscription, $);
            }

            // Process change notification
            return await this.processChange(notification, subscription, $);
          } catch (error) {
            console.error("Error processing notification:", error);
            return null;
          }
        })
      );

      // Filter out failed notifications
      const processedResults = results.filter(Boolean);

      // Report health status
      await this.reportHealth("healthy", {
        action: "notifications_processed",
        count: processedResults.length,
        total: body.value.length
      });

      // Emit processed notifications
      if (processedResults.length > 0) {
        $.export("$summary", `Processed ${processedResults.length} notifications`);
        $.export("notifications", processedResults);
      }

      // Send response to Microsoft Graph
      $.respond({
        status: 202,
        body: {
          processed: processedResults.length,
          timestamp: new Date().toISOString()
        }
      });

    } catch (error) {
      await this.handleError($, error, {
        action: "notification_processing",
        requestId: headers["request-id"]
      });

      // Always respond with success to avoid notification retries
      $.respond({
        status: 202
      });
    }
  }
}; 