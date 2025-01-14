import { axios } from "@pipedream/platform";

/**
 * Microsoft Graph API App for Pipedream
 * Handles core functionality for Microsoft Graph API integration
 */
export default {
  type: "app",
  app: "microsoft_graph",
  
  propDefinitions: {
    // Generic prop definitions that could be used across different Graph API features
    resource: {
      type: "string",
      label: "Resource",
      description: "The Microsoft Graph resource to interact with",
      options: [
        { label: "Inbox Messages", value: "me/mailFolders('Inbox')/messages" },
        { label: "Calendar Events", value: "me/events" },
        { label: "Contacts", value: "me/contacts" }
      ],
      default: "me/mailFolders('Inbox')/messages"
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
    _baseUrl() {
      return "https://graph.microsoft.com/v1.0";
    },

    /**
     * Enhanced request handling with backoff and rate limiting
     */
    async _makeRequest(opts = {}) {
      const {
        $,
        path,
        maxRetries = this.maxRetries || 3,
        ...otherOpts
      } = opts;

      const makeAttempt = async (retryCount = 0) => {
        try {
          const config = {
            url: `${this._baseUrl()}/${path}`,
            headers: {
              Authorization: `Bearer ${this.$auth.oauth_access_token}`,
              ...otherOpts?.headers,
            },
            ...otherOpts,
          };

          return await axios($, config);
        } catch (error) {
          if (retryCount >= maxRetries || !this._isRetryableError(error)) {
            throw error;
          }

          const delay = Math.min(1000 * Math.pow(2, retryCount), 10000);
          await new Promise(resolve => setTimeout(resolve, delay));
          return makeAttempt(retryCount + 1);
        }
      };

      return makeAttempt();
    },

    _isRetryableError(error) {
      return error.response?.status >= 500 || 
             error.response?.status === 429 ||
             error.code === 'ECONNRESET';
    },

    /**
     * Generic method to validate Graph API permissions
     */
    async validatePermissions($, permission, resource) {
      try {
        await this._makeRequest({
          $,
          path: resource,
          maxRetries: 1,
        });
      } catch (error) {
        if (error.response?.status === 401 || error.response?.status === 403) {
          throw new Error(`${permission} permission not granted: ${error.message}`);
        }
        throw error;
      }
    },

    /**
     * Messages API Methods
     */
    async getMessage($, messageId) {
      return this._makeRequest({
        $,
        path: `me/messages/${messageId}`,
      });
    },

    async listMessages($, params = {}) {
      return this._makeRequest({
        $,
        path: "me/messages",
        params,
      });
    },

    /**
     * Calendar API Methods
     */
    async getEvent($, eventId) {
      return this._makeRequest({
        $,
        path: `me/events/${eventId}`,
      });
    },

    async listEvents($, params = {}) {
      return this._makeRequest({
        $,
        path: "me/events",
        params,
      });
    },

    /**
     * Contacts API Methods
     */
    async getContact($, contactId) {
      return this._makeRequest({
        $,
        path: `me/contacts/${contactId}`,
      });
    },

    async listContacts($, params = {}) {
      return this._makeRequest({
        $,
        path: "me/contacts",
        params,
      });
    },

    /**
     * Subscription API Methods
     * These are kept in the app file as they're core API operations,
     * but subscription-specific logic is moved to common/subscription.mjs
     */
    async createSubscription($, {
      resource,
      webhookUrl,
      expirationDateTime,
      clientState,
      includeResourceData = false,
    }) {
      return this._makeRequest({
        $,
        method: "POST",
        path: "subscriptions",
        data: {
          changeType: "created,updated",
          notificationUrl: webhookUrl,
          resource,
          expirationDateTime,
          clientState,
          lifecycleNotificationUrl: webhookUrl,
          includeResourceData,
        },
      });
    },

    async renewSubscription($, subscriptionId, expirationDateTime) {
      return this._makeRequest({
        $,
        method: "PATCH",
        path: `subscriptions/${subscriptionId}`,
        data: {
          expirationDateTime,
        },
      });
    },

    async getSubscription($, subscriptionId) {
      if (!subscriptionId) {
        throw new Error("Subscription ID is required");
      }

      return this._makeRequest({
        $,
        path: `subscriptions/${subscriptionId}`,
      });
    },

    async deleteSubscription($, subscriptionId) {
      return this._makeRequest({
        $,
        method: "DELETE",
        path: `subscriptions/${subscriptionId}`,
      });
    }
  },
}; 