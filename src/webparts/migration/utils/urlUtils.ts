/**
 * Utility functions for handling SharePoint URLs reliably.
 */

/**
 * Normalizes a SharePoint URL to prevent double-protocol or double-domain issues.
 * If the provided URL is already absolute, it returns the URL as-is.
 * If it's relative, it safely prepends the web absolute URL.
 * 
 * @param webUrl The base SharePoint web absolute URL (e.g., https://tenant.sharepoint.com/sites/site)
 * @param path The path to normalize (could be absolute or relative)
 * @returns A safe, absolute SharePoint URL
 */
export const normalizeSharePointUrl = (webUrl: string, path: string): string => {
  if (!path) return webUrl;

  const cleanPath = path.trim();
  const cleanWebUrl = webUrl.trim().replace(/\/$/, '');

  // Case 1: path is already an absolute URL
  if (/^https?:\/\//i.test(cleanPath)) {
    return cleanPath;
  }

  // Case 2: path is server-relative (starts with /)
  // We need to check if the path already includes the site collection path
  // If we just blindly prepend webUrl, we might get site/site/library
  if (cleanPath.startsWith('/')) {
    try {
      const url = new URL(cleanWebUrl);
      const origin = url.origin;
      return `${origin}${cleanPath}`;
    } catch {
      // Fallback if webUrl isn't a valid URL for some reason
      return `${cleanWebUrl}${cleanPath}`;
    }
  }

  // Case 3: path is relative to the current site (doesn't start with / or protocol)
  return `${cleanWebUrl}/${cleanPath}`;
};

/**
 * Strips the domain and protocol from a URL to make it server-relative.
 * Useful for GetFileByServerRelativeUrl API calls.
 */
export const makeServerRelativeUrl = (url: string): string => {
  if (!url) return '';
  
  // If it's already server-relative, just return it
  if (url.startsWith('/') && !url.startsWith('//')) {
    return url;
  }

  try {
    const parsed = new URL(url);
    return parsed.pathname;
  } catch {
    // If not a valid absolute URL, return as is
    return url;
  }
};
