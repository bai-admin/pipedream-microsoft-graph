/**
 * Common subscription functionality for Microsoft Graph components
 * Provides shared logic and prop definitions for subscription management
 */

export const SUBSCRIPTION_PROPS = {
  webhookUrl: {
    type: "string",
    label: "Webhook URL",
    description: "HTTPS URL where notifications should be sent",
    secret: true
  },
  includeResourceData: {
    type: "boolean",
    label: "Include Resource Data",
    description: "Whether to include resource data in notifications (requires additional setup)",
    optional: true,
    default: false
  },
  renewalThresholdHours: {
    type: "integer",
    label: "Renewal Threshold",
    description: "Hours before expiration when subscription should be renewed",
    optional: true,
    default: 12,
    min: 1,
    max: 24
  },
  dataRetentionDays: {
    type: "integer",
    label: "Data Retention Period",
    description: "Number of days to retain subscription data",
    optional: true,
    default: 30,
    min: 1,
    max: 90
  }
};

export const SUBSCRIPTION_METHODS = {
  /**
   * Validates subscription configuration
   */
  async validateConfig($, webhookUrl) {
    if (!webhookUrl) {
      throw new Error("Webhook URL is required");
    }
    
    if (!webhookUrl.startsWith("https://")) {
      throw new Error("Webhook URL must use HTTPS protocol");
    }
    
    await this.microsoft_graph.validatePermissions($, "Mail.Read", "me/messages");
  },

  /**
   * Retrieves stored subscription metadata
   */
  async getStoredSubscription() {
    const subscription = await this.data_store.get("subscription");
    if (!subscription) {
      throw new Error("No subscription found in data store");
    }
    return subscription;
  },

  /**
   * Stores subscription metadata with enhanced attributes
   */
  async storeSubscription(subscription) {
    const metadata = {
      id: subscription.id,
      resource: subscription.resource,
      expirationDateTime: subscription.expirationDateTime,
      clientState: subscription.clientState,
      createdDateTime: subscription.createdDateTime || new Date().toISOString(),
      lastRenewalAttempt: new Date().toISOString(),
      status: "active",
      webhookUrl: subscription.notificationUrl,
      includeResourceData: subscription.includeResourceData
    };

    await this.data_store.set("subscription", metadata);
    return metadata;
  },

  /**
   * Reports component health status with detailed information
   */
  async reportHealth(status, details = {}) {
    const healthStatus = {
      status,
      timestamp: new Date().toISOString(),
      details: {
        ...details,
        component: this.key
      }
    };

    await this.data_store.set("health_status", healthStatus);
    console.log("Health Status:", JSON.stringify(healthStatus, null, 2));
    return healthStatus;
  },

  /**
   * Enhanced error handling with context
   */
  async handleError($, error, context = {}) {
    const errorDetails = {
      message: error.message,
      code: error.code,
      context,
      timestamp: new Date().toISOString()
    };

    if (error.response) {
      errorDetails.status = error.response.status;
      errorDetails.statusText = error.response.statusText;
      errorDetails.data = error.response.data;
    }

    await this.reportHealth("error", errorDetails);
    console.error("Error occurred:", JSON.stringify(errorDetails, null, 2));
    
    if ($ && $.flow) {
      $.flow.exit(errorDetails.message);
    } else {
      throw error;
    }
  },

  /**
   * Calculates time until subscription expiration
   */
  getTimeUntilExpiration(expirationDateTime) {
    const now = new Date();
    const expiration = new Date(expirationDateTime);
    return Math.floor((expiration - now) / (1000 * 60)); // minutes
  },

  /**
   * Cleans up old subscription data
   */
  async cleanupOldData(retentionDays = 30) {
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - retentionDays);

    const keys = await this.data_store.keys();
    const batchSize = 100;
    
    for (let i = 0; i < keys.length; i += batchSize) {
      const batch = keys.slice(i, i + batchSize);
      for (const key of batch) {
        const data = await this.data_store.get(key);
        if (data?.timestamp && new Date(data.timestamp) < cutoffDate) {
          await this.data_store.delete(key);
        }
      }
    }
  },

  /**
   * Acquires a processing lock
   */
  async acquireLock(resourceId, ttlSeconds = 60) {
    const lockKey = `lock:${resourceId}`;
    const now = Date.now();
    const lock = await this.data_store.get(lockKey);
    
    if (lock && now < lock.expiresAt) {
      return false;
    }
    
    await this.data_store.set(lockKey, {
      acquiredAt: now,
      expiresAt: now + (ttlSeconds * 1000)
    });
    
    return true;
  },

  /**
   * Releases a processing lock
   */
  async releaseLock(resourceId) {
    const lockKey = `lock:${resourceId}`;
    await this.data_store.delete(lockKey);
  }
}; 