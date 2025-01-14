import common from "../../common/subscription.mjs";

/**
 * Source component that provides Microsoft Graph subscription metadata
 */
export default {
  ...common,
  
  key: "microsoft_graph-subscription-metadata",
  name: "Microsoft Graph Subscription Metadata",
  description: "Provides status information about Microsoft Graph API subscriptions",
  version: "0.0.1",
  type: "source",
  
  props: {
    ...common.props,
    timer: {
      type: "$.interface.timer",
      default: {
        intervalSeconds: 60 * 60, // Run every hour
      },
    }
  },

  async run({ steps, $ }) {
    try {
      // Get current subscription info
      const subscription = await this.getStoredSubscription();
      
      if (!subscription) {
        throw new Error("No subscription found in data store");
      }

      // Calculate time until expiration
      const timeUntilExpiration = this.getTimeUntilExpiration(subscription.expirationDateTime);
      
      // Get current subscription status from Microsoft Graph
      const currentStatus = await this.microsoft_graph.getSubscription($, subscription.id);
      
      // Prepare metadata
      const metadata = {
        subscriptionId: subscription.id,
        timeUntilExpiration, // minutes
        lastRenewalAttempt: subscription.lastRenewalAttempt,
        resource: subscription.resource,
        status: currentStatus.status,
        created: subscription.created
      };

      // Report health status
      await this.reportHealth("healthy", {
        message: "Subscription metadata retrieved successfully",
        ...metadata
      });

      // Emit the metadata
      this.$emit(metadata, {
        summary: `Subscription status for ${subscription.id}`,
        ts: Date.now(),
      });

      return metadata;
    } catch (error) {
      await this.handleError($, error, "subscription_metadata");
      throw error;
    }
  },
}; 