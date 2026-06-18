import { SPHttpClient } from '@microsoft/sp-http';

type TDownloadContext = {
  spHttpClient: SPHttpClient;
};

const escapeODataString = (value: string): string => value.replace(/'/g, "''");

const toServerRelativeUrl = (fileUrl: string, siteUrl?: string): string => {
  const trimmed = (fileUrl || '').trim();
  if (!trimmed) {
    return '';
  }

  if (trimmed.charAt(0) === '/') {
    return decodeURIComponent(trimmed);
  }

  try {
    return decodeURIComponent(new URL(trimmed).pathname);
  } catch {
    if (siteUrl && trimmed.toLowerCase().indexOf(siteUrl.toLowerCase()) === 0) {
      return decodeURIComponent(trimmed.slice(siteUrl.length));
    }
  }

  return trimmed.charAt(0) === '/' ? decodeURIComponent(trimmed) : `/${decodeURIComponent(trimmed)}`;
};

const triggerBrowserDownload = async (blob: Blob, fileName: string): Promise<void> => {
  const objectUrl = window.URL.createObjectURL(blob);
  const link = window.document.createElement('a');
  link.href = objectUrl;
  link.download = fileName || 'download';
  link.rel = 'noopener noreferrer';
  link.style.display = 'none';
  window.document.body.appendChild(link);
  link.click();
  window.document.body.removeChild(link);
  window.setTimeout(() => window.URL.revokeObjectURL(objectUrl), 1000);
};

const triggerNativeSharePointDownload = (webUrl: string, serverRelativeUrl: string): void => {
  const downloadUrl =
    `${webUrl}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(serverRelativeUrl)}`;
  const link = window.document.createElement('a');
  link.href = downloadUrl;
  link.rel = 'noopener noreferrer';
  link.style.display = 'none';
  window.document.body.appendChild(link);
  link.click();
  window.document.body.removeChild(link);
};

const openDirectDownloadFallback = (webUrl: string, serverRelativeUrl: string): void => {
  const directUrl = `${webUrl}${serverRelativeUrl}`;
  const separator = directUrl.indexOf('?') === -1 ? '?' : '&';
  window.open(`${directUrl}${separator}download=1`, '_blank', 'noopener,noreferrer');
};

export const downloadSharePointFile = async (
  context: TDownloadContext,
  webUrl: string,
  fileUrlOrServerRelativeUrl: string,
  fileName: string
): Promise<void> => {
  const serverRelativeUrl = toServerRelativeUrl(fileUrlOrServerRelativeUrl, webUrl);
  if (!serverRelativeUrl) {
    throw new Error('File URL is not available.');
  }

  try {
    triggerNativeSharePointDownload(webUrl, serverRelativeUrl);
    return;
  } catch {
    // Fall back to the authenticated blob path below.
  }

  const encodedServerRelativeUrl = escapeODataString(serverRelativeUrl);
  const downloadUrl =
    `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${encodedServerRelativeUrl}')/$value`;

  const response = await context.spHttpClient.get(
    downloadUrl,
    SPHttpClient.configurations.v1
  );

  if (!response.ok) {
    openDirectDownloadFallback(webUrl, serverRelativeUrl);
    return;
  }

  const blob = await response.blob();
  if (!blob || blob.size === 0) {
    openDirectDownloadFallback(webUrl, serverRelativeUrl);
    return;
  }

  await triggerBrowserDownload(blob, fileName);
};
