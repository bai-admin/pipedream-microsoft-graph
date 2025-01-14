import crypto from "crypto";
import { SUBSCRIPTION_PROPS, SUBSCRIPTION_METHODS } from "../../common/subscription.mjs";

/**
 * Source component that creates and manages Microsoft Graph subscriptions
 */
export default {
  key: "microsoft_graph-subscription-creator",
  name: "Microsoft Graph Subscription Creator",
  description: "Creates and manages subscriptions to Microsoft Graph resources",
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
    resource: {
      type: "string",
      label: "Resource",
      description: "The resource to subscribe to",
      options: [
        { label: "Inbox Messages", value: "me/mailFolders('Inbox')/messages" },
        { label: "Calendar Events", value: "me/events" },
        { label: "Contacts", value: "me/contacts" }
      ],
      default: "me/mailFolders('Inbox')/messages"
    }
  },

  methods: {
    ...SUBSCRIPTION_METHODS,

    /**
     * Generates a unique client state for subscription validation
     */
    generateClientState() {
      return crypto.randomBytes(16).toString("hex");
    },

    /**
     * Calculates subscription expiration date (max 3 days from now)
     */
    getExpirationDateTime() {
      const now = new Date();
      const expiration = new Date(now);
      expiration.setDate(expiration.getDate() + 3);
      return expiration.toISOString();
    }
  },

  async deploy() {
    await this.validateConfig(this.$, this.webhookUrl);
  },

  async run({ $ }) {
    try {
      // Generate subscription parameters
      const expirationDateTime = this.getExpirationDateTime();
      const clientState = this.generateClientState();

      // Create subscription
      const subscription = await this.microsoft_graph.createSubscription($, {
        resource: this.resource,
        webhookUrl: this.webhookUrl,
        expirationDateTime,
        clientState,
        includeResourceData: this.includeResourceData
      });

      // Store subscription metadata
      const metadata = await this.storeSubscription(subscription);

      // Report successful creation
      await this.reportHealth("healthy", {
        action: "subscription_created",
        subscriptionId: subscription.id,
        expirationDateTime
      });

      // Emit subscription for downstream processing
      $.export("$summary", `Created subscription for ${this.resource}`);
      $.export("subscription", metadata);

    } catch (error) {
      await this.handleError($, error, {
        action: "subscription_creation",
        resource: this.resource
      });
    }
  }
}; 