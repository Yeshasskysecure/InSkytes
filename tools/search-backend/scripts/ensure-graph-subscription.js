const { loadEnv, requireValue, toNumber } = require('./lib/env');
const {
  createSubscription,
  deleteSubscription,
  getGraphToken,
  listSubscriptions,
  updateSubscription
} = require('./lib/graphClient');
const { withRetry } = require('./lib/retry');

const envPath = process.argv[2] || 'config/search.env';
const action = process.argv[3] || 'ensure';
const confirmedDelete = process.argv.includes('--yes-delete-subscription');

const envValue = (env, key) => env[key] || process.env[key];
const requireEnvValue = (env, key) => envValue(env, key) || requireValue(env, key);
const normalizeResource = (value) => String(value || '').replace(/^\/+/, '').toLowerCase();

const subscriptionResource = (env) =>
  String(envValue(env, 'GRAPH_WEBHOOK_RESOURCE') || `drives/${requireValue(env, 'SHAREPOINT_LIBRARY_DRIVE_ID')}/root`).replace(/^\/+/, '');

const subscriptionExpiration = (env) => {
  const hours = Math.min(Math.max(toNumber(envValue(env, 'GRAPH_WEBHOOK_EXPIRATION_HOURS'), 48), 1), 168);
  return new Date(Date.now() + hours * 60 * 60 * 1000).toISOString();
};

const findMatchingSubscriptions = (subscriptions, resource, notificationUrl) => {
  const normalizedResource = normalizeResource(resource);
  const normalizedUrl = String(notificationUrl || '').toLowerCase();
  return subscriptions.filter((subscription) =>
    normalizeResource(subscription.resource) === normalizedResource &&
    String(subscription.notificationUrl || '').toLowerCase() === normalizedUrl
  );
};

const run = async () => {
  const { env, absolutePath } = loadEnv(envPath);
  const token = await getGraphToken(env);
  const notificationUrl = requireEnvValue(env, 'GRAPH_WEBHOOK_NOTIFICATION_URL');
  const clientState = requireEnvValue(env, 'GRAPH_WEBHOOK_CLIENT_STATE');
  const resource = subscriptionResource(env);
  const changeType = envValue(env, 'GRAPH_WEBHOOK_CHANGE_TYPE') || 'updated';
  const expirationDateTime = subscriptionExpiration(env);

  console.log(`Using env: ${absolutePath}`);
  console.log(`Webhook resource: ${resource}`);
  console.log(`Webhook notification URL: ${notificationUrl}`);
  console.log(`Action: ${action}`);

  const current = await withRetry(
    () => listSubscriptions(token),
    { retries: 4, baseDelayMs: 1500, label: 'list Graph subscriptions' }
  );
  const matches = findMatchingSubscriptions(current.value || [], resource, notificationUrl);

  if (action === 'list') {
    console.log(JSON.stringify({
      total: (current.value || []).length,
      matches: matches.map((subscription) => ({
        id: subscription.id,
        resource: subscription.resource,
        changeType: subscription.changeType,
        expirationDateTime: subscription.expirationDateTime,
        notificationUrl: subscription.notificationUrl
      }))
    }, null, 2));
    return;
  }

  if (action === 'delete') {
    if (!confirmedDelete) {
      throw new Error('Refusing to delete Graph webhook subscriptions without --yes-delete-subscription.');
    }
    for (const subscription of matches) {
      await withRetry(
        () => deleteSubscription(token, subscription.id),
        { retries: 4, baseDelayMs: 1500, label: `delete subscription ${subscription.id}` }
      );
    }
    console.log(JSON.stringify({ deleted: matches.length }, null, 2));
    return;
  }

  if (matches.length > 0) {
    const [subscription] = matches;
    const updated = await withRetry(
      () => updateSubscription(token, subscription.id, { expirationDateTime }),
      { retries: 4, baseDelayMs: 1500, label: `renew subscription ${subscription.id}` }
    );
    console.log(JSON.stringify({
      action: 'renewed',
      id: updated.id,
      resource: updated.resource,
      expirationDateTime: updated.expirationDateTime,
      duplicateMatches: Math.max(0, matches.length - 1)
    }, null, 2));
    return;
  }

  const created = await withRetry(
    () => createSubscription(token, {
      changeType,
      notificationUrl,
      resource,
      expirationDateTime,
      clientState
    }),
    { retries: 4, baseDelayMs: 1500, label: 'create Graph subscription' }
  );
  console.log(JSON.stringify({
    action: 'created',
    id: created.id,
    resource: created.resource,
    expirationDateTime: created.expirationDateTime
  }, null, 2));
};

run().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
