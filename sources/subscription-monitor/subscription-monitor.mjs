import { SUBSCRIPTION_PROPS, SUBSCRIPTION_METHODS } from "../../common/subscription.mjs";

/**
 * Source component that monitors and renews Microsoft Graph subscriptions
 */
export default {
  key: "microsoft_graph-subscription-monitor",
  name: "Microsoft Graph Subscription Monitor",
  description: "Monitors and automatically renews Microsoft Graph subscriptions before expiration",
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
    timer: {
      type: "$.interface.timer",
      label: "Monitoring Schedule",
      description: "How often to check subscription status",
      default: {
        intervalSeconds: 60 * 60 * 6, // Every 6 hours
      },
    }
  },

  methods: {
    ...SUBSCRIPTION_METHODS,

    /**
     * Calculates new expiration date for renewal (max 3 days)
     */
    getNewExpirationDateTime() {
      const now = new Date();
      const expiration = new Date(now);
      expiration.setDate(expiration.getDate() + 3);
      return expiration.toISOString();
    },

    /**
     * Renews subscription with enhanced error handling
     */
    async renewSubscription($, subscription) {
      const expirationDateTime = this.getNewExpirationDateTime();
      
      const renewedSubscription = await this.microsoft_graph.renewSubscription($, 
        subscription.id, 
        expirationDateTime
      );

      const metadata = await this.storeSubscription({
        ...renewedSubscription,
        notificationUrl: subscription.webhookUrl,
        includeResourceData: subscription.includeResourceData
      });

      await this.reportHealth("healthy", {
        action: "subscription_renewed",
        subscriptionId: subscription.id,
        expirationDateTime
      });

      return metadata;
    }
  },

  async run({ $ }) {
    try {
      // Clean up old data based on retention period
      await this.cleanupOldData(this.dataRetentionDays);

      // Get current subscription
      const subscription = await this.getStoredSubscription();
      
      // Check time until expiration
      const minutesUntilExpiration = this.getTimeUntilExpiration(subscription.expirationDateTime);
      const renewalThresholdMinutes = this.renewalThresholdHours * 60;

      // Renew if within threshold
      if (minutesUntilExpiration <= renewalThresholdMinutes) {
        // Acquire lock to prevent concurrent renewals
        const hasLock = await this.acquireLock(subscription.id);
        if (!hasLock) {
          console.log("Another instance is handling renewal");
          return;
        }

        try {
          const renewedSubscription = await this.renewSubscription($, subscription);
          $.export("$summary", `Renewed subscription ${subscription.id}`);
          $.export("subscription", renewedSubscription);
        } finally {
          await this.releaseLock(subscription.id);
        }
      } else {
        await this.reportHealth("healthy", {
          action: "subscription_checked",
          subscriptionId: subscription.id,
          minutesUntilExpiration
        });
        
        $.export("$summary", `Subscription valid for ${Math.floor(minutesUntilExpiration / 60)} hours`);
      }
    } catch (error) {
      await this.handleError($, error, {
        action: "subscription_monitoring"
      });
    }
  }
}; 