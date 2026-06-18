import { MSGraphClientV3 } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { CACHE_KEYS, GROUP_IDS, GROUP_NAMES } from '../config/appConfig';

export interface IKmsUser {
  displayName: string;
  upn: string;
  email: string;
}

const KMS_USERS_GROUP_ID = GROUP_IDS.kmsUsers;
const GROUP_CACHE_KEY = CACHE_KEYS.kmsUsersGroupId;
const KMS_USERS_CACHE_KEY = CACHE_KEYS.kmsUsers;
const KMS_USERS_CACHE_TTL = 30 * 60 * 1000;

let _inMemoryUsersCache: IKmsUser[] | null = null;

const getCachedGroupId = (): string => {
  try {
    return sessionStorage.getItem(GROUP_CACHE_KEY) || localStorage.getItem(GROUP_CACHE_KEY) || '';
  } catch (error) {
    return '';
  }
};

const cacheGroupId = (groupId: string): void => {
  if (!groupId) {
    return;
  }

  try {
    sessionStorage.setItem(GROUP_CACHE_KEY, groupId);
    localStorage.setItem(GROUP_CACHE_KEY, groupId);
  } catch (error) {
    // Storage can fail in private mode or restricted contexts. Ignore safely.
  }
};

const getCachedKmsUsers = (): IKmsUser[] | null => {
  try {
    const cached = sessionStorage.getItem(KMS_USERS_CACHE_KEY);
    if (!cached) return null;

    const parsed = JSON.parse(cached);
    if (Date.now() - parsed.timestamp > KMS_USERS_CACHE_TTL) return null;

    return parsed.users || null;
  } catch (error) {
    return null;
  }
};

const setCachedKmsUsers = (users: IKmsUser[]): void => {
  try {
    sessionStorage.setItem(KMS_USERS_CACHE_KEY, JSON.stringify({
      timestamp: Date.now(),
      users
    }));
  } catch (error) {
    // Storage can fail in private mode or restricted contexts. Ignore safely.
  }
};

export const clearKmsUsersCache = (): void => {
  _inMemoryUsersCache = null;
  try {
    sessionStorage.removeItem('kmKmsUsers_v1');
  } catch (error) {
    // Storage can fail in private mode or restricted contexts. Ignore safely.
  }
};

const getGraphClient = async (context: WebPartContext): Promise<MSGraphClientV3> =>
  context.msGraphClientFactory.getClient('3');

const getRelativeGraphPath = (pathOrUrl: string): string =>
  pathOrUrl.indexOf('https://graph.microsoft.com') === 0
    ? pathOrUrl
      .replace('https://graph.microsoft.com/v1.0', '')
      .replace('https://graph.microsoft.com/v1.0', '')
      .replace('https://graph.microsoft.com', '')
      .replace(/^\/v1\.0/i, '')
    : pathOrUrl.replace(/^\/v1\.0/i, '');

export const getKmsUsersGroupId = async (context: WebPartContext): Promise<string> => {
  const cachedGroupId = getCachedGroupId();
  if (cachedGroupId) {
    return cachedGroupId;
  }

  if (KMS_USERS_GROUP_ID) {
    cacheGroupId(KMS_USERS_GROUP_ID);
    return KMS_USERS_GROUP_ID;
  }

  const client = await getGraphClient(context);
  const response = await client
    .api(`/groups?$filter=displayName eq '${GROUP_NAMES.kmsUsers}'&$select=id,displayName`)
    .get();

  const groupId = (response?.value || [])[0]?.id || '';
  if (!groupId) {
    throw new Error(`Entra group "${GROUP_NAMES.kmsUsers}" was not found.`);
  }

  cacheGroupId(groupId);
  return groupId;
};

export const getKmsUsers = async (context: WebPartContext): Promise<IKmsUser[]> => {
  if (_inMemoryUsersCache) {
    return _inMemoryUsersCache;
  }

  const cachedUsers = getCachedKmsUsers();
  if (cachedUsers) {
    _inMemoryUsersCache = cachedUsers;
    return cachedUsers;
  }

  const client = await getGraphClient(context);
  const groupId = await getKmsUsersGroupId(context);
  let requestPath = `/groups/${groupId}/members/microsoft.graph.user?$select=id,displayName,mail,userPrincipalName&$top=999`;
  const usersByUpn: Record<string, IKmsUser> = {};

  while (requestPath) {
    const response = await client.api(getRelativeGraphPath(requestPath)).version('v1.0').get();
    const members = response?.value || [];

    for (let index = 0; index < members.length; index++) {
      const member = members[index];
      const upn = member?.userPrincipalName || '';
      const displayName = member?.displayName || upn;

      if (!upn || !displayName) {
        continue;
      }

      usersByUpn[upn.toLowerCase()] = {
        displayName,
        upn,
        email: member?.mail || upn
      };
    }

    requestPath = response?.['@odata.nextLink'] || '';
  }

  const users = Object.keys(usersByUpn)
    .map((key) => usersByUpn[key])
    .sort((left, right) => left.displayName.localeCompare(right.displayName));

  _inMemoryUsersCache = users;
  setCachedKmsUsers(users);

  return users;
};
