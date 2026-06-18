import { override } from '@microsoft/decorators';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  IListViewCommandSetListViewUpdatedParameters
} from '@microsoft/sp-listview-extensibility';
import { LIBRARY_NAMES, LIST_NAMES } from '../../webparts/migration/config/appConfig';
 
const SYNC_COMMAND_ID = 'SYNC_METRICS';
const KM_DATA_HUB_LIBRARY = LIBRARY_NAMES.kmDataHub;
const SYNC_CONFIG_LIST = LIST_NAMES.syncConfig;
const DOCUMENT_METRICS_LIST = LIST_NAMES.documentMetrics;

const isDebugLoggingEnabled = (): boolean => {
  try {
    return typeof window !== 'undefined' &&
      window.localStorage?.getItem('IKNOWLEDGE_DEBUG_LOGS') === 'true';
  } catch {
    return false;
  }
};

const debugLog = (...args: unknown[]): void => {
  if (isDebugLoggingEnabled()) {
    console.log(...args);
  }
};

type TSyncHttpResponse = {
  ok: boolean;
  status: number;
  headers: Headers;
  json: () => Promise<any>;
};

export interface ISyncMetricsCommandSetProperties {
}
 
export default class SyncMetricsCommandSet extends BaseListViewCommandSet<ISyncMetricsCommandSetProperties> {
  private _canSyncMetrics: boolean = false;
  private siteUrl: string = '';
 
  private get spHttpClientConfiguration(): any {
    return (this.context.spHttpClient as any).constructor.configurations.v1;
  }
 
  @override
  public async onInit(): Promise<void> {
    this.siteUrl = this.context.pageContext.web.absoluteUrl;
    this._canSyncMetrics = await this.isAdmin();
  }
 
  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const syncCommand: Command = this.tryGetCommand(SYNC_COMMAND_ID);
    if (syncCommand) {
      syncCommand.visible =
        this._canSyncMetrics &&
        this.context.pageContext.list?.title === KM_DATA_HUB_LIBRARY;
    }
  }
 
  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    if (event.itemId === SYNC_COMMAND_ID) {
      this.syncPendingMetrics().catch(err => {
        console.error('[SYNC] Fatal:', err);
        this.showToast('Sync failed - see console');
      });
    }
  }
 
  private async isAdmin(): Promise<boolean> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      const ownerResp = await this.context.spHttpClient.get(
        `${webUrl}/_api/web/DoesUserHavePermissions(@v)?@v={'High':'0','Low':'67108864'}`,
        this.spHttpClientConfiguration,
        { headers: { Accept: 'application/json;odata=nometadata' } }
      );
      if (ownerResp.ok) {
        const ownerData = await ownerResp.json();
        if (ownerData?.value === true) {
          return true;
        }
      }
 
      const resp = await this.context.spHttpClient.get(
        `${webUrl}/_api/web/currentuser/isSiteAdmin`,
        this.spHttpClientConfiguration,
        { headers: { Accept: 'application/json;odata=nometadata' } }
      );
      if (!resp.ok) return false;
      const data = await resp.json();
      if (data?.value === true) {
        return true;
      }
    } catch {
      return false;
    }
 
    return false;
  }
 
  private async ensureSyncConfigList(): Promise<boolean> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
 
    try {
      const existingResp = await this.context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${SYNC_CONFIG_LIST}')?$select=Id`,
        this.spHttpClientConfiguration,
        { headers: { Accept: 'application/json;odata=nometadata' } }
      );
 
      if (existingResp.ok) {
        return true;
      }
 
      const createResp = await this.context.spHttpClient.post(
        `${webUrl}/_api/web/lists`,
        this.spHttpClientConfiguration,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata'
          },
          body: JSON.stringify({
            Title: SYNC_CONFIG_LIST,
            BaseTemplate: 100
          })
        }
      );
 
      if (!createResp.ok) {
        console.warn(`Unable to create ${SYNC_CONFIG_LIST}: ${createResp.status}`);
        return false;
      }
 
      const fieldResp = await this.context.spHttpClient.post(
        `${webUrl}/_api/web/lists/getbytitle('${SYNC_CONFIG_LIST}')/fields/createfieldasxml`,
        this.spHttpClientConfiguration,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata'
          },
          body: JSON.stringify({
            parameters: {
              SchemaXml: '<Field DisplayName="ConfigValue" Name="ConfigValue" Type="Text" />',
              Options: 0
            }
          })
        }
      );
 
      if (!fieldResp.ok) {
        console.warn(`Unable to create ConfigValue field in ${SYNC_CONFIG_LIST}: ${fieldResp.status}`);
        return false;
      }
 
      return true;
    } catch (err) {
      console.warn('ensureSyncConfigList error:', err);
      return false;
    }
  }
 
  private async getLastSyncTime(): Promise<Date | null> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      const resp = await this.context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${SYNC_CONFIG_LIST}')/items` +
        `?$select=Id,ConfigValue&$filter=Title eq 'LastMetricsSyncTime'&$top=1`,
        this.spHttpClientConfiguration,
        { headers: { Accept: 'application/json;odata=nometadata' } }
      );
      if (!resp.ok) return null;
      const data = await resp.json();
      const items = data?.value || [];
      if (items.length === 0) return null;
      return new Date(items[0].ConfigValue);
    } catch {
      return null;
    }
  }
 
  private async updateLastSyncTime(): Promise<void> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      const now = new Date().toISOString();
 
      // Check if record exists
      const resp = await this.context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${SYNC_CONFIG_LIST}')/items` +
        `?$select=Id&$filter=Title eq 'LastMetricsSyncTime'&$top=1`,
        this.spHttpClientConfiguration,
        { headers: { Accept: 'application/json;odata=nometadata' } }
      );
      const data = await resp.json();
      const items = data?.value || [];
 
      if (items.length > 0) {
        // Update existing
        await this.context.spHttpClient.post(
          `${webUrl}/_api/web/lists/getbytitle('${SYNC_CONFIG_LIST}')/items(${items[0].Id})`,
          this.spHttpClientConfiguration,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-Type': 'application/json;odata=nometadata',
              'X-HTTP-Method': 'MERGE',
              'IF-MATCH': '*'
            },
            body: JSON.stringify({ ConfigValue: now })
          }
        );
      } else {
        // Create new
        await this.context.spHttpClient.post(
          `${webUrl}/_api/web/lists/getbytitle('${SYNC_CONFIG_LIST}')/items`,
          this.spHttpClientConfiguration,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-Type': 'application/json;odata=nometadata'
            },
            body: JSON.stringify({
              Title: 'LastMetricsSyncTime',
              ConfigValue: now
            })
          }
        );
      }
    } catch (err) {
      console.warn('updateLastSyncTime error:', err);
    }
  }
 
  private async retryRequest(
    fn: () => Promise<TSyncHttpResponse>,
    maxRetries: number = 3
  ): Promise<TSyncHttpResponse> {
    let lastError: any;
 
    for (let attempt = 0; attempt < maxRetries; attempt++) {
      try {
        const resp = await fn();
        if (resp.status === 429 || resp.status === 503) {
          const retryAfter = parseInt(resp.headers.get('Retry-After') || '10', 10);
          console.warn(`[SYNC] Throttled, retrying in ${retryAfter}s`);
          await new Promise(resolve => setTimeout(resolve, retryAfter * 1000));
          continue;
        }
 
        return resp;
      } catch (err) {
        lastError = err;
        await new Promise(resolve => setTimeout(resolve, 2000 * (attempt + 1)));
      }
    }
 
    throw lastError;
  }
 
  private encodeFileRef(fileRef: string): string {
    try {
      const decoded = decodeURIComponent(fileRef);
      return encodeURIComponent(decoded);
    } catch {
      return encodeURIComponent(fileRef);
    }
  }
 
  private async getFreshDigest(): Promise<string> {
    const resp = await this.context.spHttpClient.post(
      `${this.siteUrl}/_api/contextinfo`,
      this.spHttpClientConfiguration,
      { body: '' }
    );
    const data = await resp.json();
    return data.d?.GetContextWebInformation?.FormDigestValue || data.FormDigestValue || '';
  }
 
  private async getSyncConfig(key: string): Promise<string> {
    try {
      const resp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${SYNC_CONFIG_LIST}')/items` +
        `?$filter=Title eq '${key}'&$select=Id,ConfigValue&$top=1`,
        this.spHttpClientConfiguration
      );
      const data = await resp.json();
      const item = data.d?.results?.[0] ?? data.value?.[0];
      return item?.ConfigValue || '';
    } catch {
      return '';
    }
  }
 
  private async setSyncConfig(key: string, value: string): Promise<void> {
    try {
      const getResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${SYNC_CONFIG_LIST}')/items` +
        `?$filter=Title eq '${key}'&$select=Id&$top=1`,
        this.spHttpClientConfiguration
      );
      const getData = await getResp.json();
      const existing = getData.d?.results?.[0] ?? getData.value?.[0];
      const digest = await this.getFreshDigest();
 
      if (existing) {
        await this.context.spHttpClient.post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${SYNC_CONFIG_LIST}')/items(${existing.Id})`,
          this.spHttpClientConfiguration,
          {
            headers: {
              'X-HTTP-Method': 'MERGE',
              'IF-MATCH': '*',
              'X-RequestDigest': digest,
              'Content-Type': 'application/json;odata=verbose',
              'Accept': 'application/json;odata=verbose',
              'OData-Version': ''
            },
            body: JSON.stringify({
              __metadata: { type: 'SP.Data.KM_x005f_SyncConfigListItem' },
              ConfigValue: value
            })
          }
        );
      } else {
        await this.context.spHttpClient.post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${SYNC_CONFIG_LIST}')/items`,
          this.spHttpClientConfiguration,
          {
            headers: {
              'X-RequestDigest': digest,
              'Content-Type': 'application/json;odata=verbose',
              'Accept': 'application/json;odata=verbose',
              'OData-Version': ''
            },
            body: JSON.stringify({
              __metadata: { type: 'SP.Data.KM_x005f_SyncConfigListItem' },
              Title: key,
              ConfigValue: value
            })
          }
        );
      }
    } catch (err) {
      console.error('[SYNC CONFIG] Error:', key, err);
    }
  }
 
  private showToast(message: string): void {
    const existing = document.getElementById('metrics-sync-toast');
    if (existing) existing.remove();
 
    const toast = document.createElement('div');
    toast.id = 'metrics-sync-toast';
    toast.style.cssText =
      'position:fixed;bottom:24px;right:24px;' +
      'background:#323130;color:#ffffff;' +
      'padding:14px 20px;border-radius:4px;' +
      'font-size:14px;font-family:Segoe UI,sans-serif;' +
      'z-index:999999;max-width:380px;' +
      'box-shadow:0 4px 16px rgba(0,0,0,0.3);' +
      'transition:opacity 0.3s ease;line-height:1.4;';
    toast.textContent = message;
    document.body.appendChild(toast);
 
    if (message.startsWith('Sync complete') || message.startsWith('Sync failed')) {
      setTimeout(() => {
        toast.style.opacity = '0';
        setTimeout(() => toast.remove(), 300);
      }, 5000);
    }
  }
 
  private async clearPendingSync(metricsItemId: number): Promise<void> {
    try {
      const digest = await this.getFreshDigest();
      await this.retryRequest(() =>
        this.context.spHttpClient.post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${DOCUMENT_METRICS_LIST}')/items(${metricsItemId})`,
          this.spHttpClientConfiguration,
          {
            headers: {
              'X-HTTP-Method': 'MERGE',
              'IF-MATCH': '*',
              'X-RequestDigest': digest,
              'Content-Type': 'application/json;odata=verbose',
              'Accept': 'application/json;odata=verbose',
              'OData-Version': ''
            },
            body: JSON.stringify({
              __metadata: { type: 'SP.Data.DocumentMetricsListItem' },
              PendingSync: false
            })
          }
        )
      );
    } catch (err) {
      console.error('[CLEAR PENDING] Error:', err);
    }
  }
 
  private async undoCheckout(encodedFileRef: string): Promise<void> {
    try {
      const digest = await this.getFreshDigest();
      await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/GetFileByServerRelativeUrl('${encodedFileRef}')/undocheckout()`,
        this.spHttpClientConfiguration,
        { headers: { 'X-RequestDigest': digest }, body: '' }
      );
    } catch (err) {
      console.error('[SYNC] Undo checkout failed:', err);
    }
  }
 
  private async syncOneToKMHub(metricItem: any): Promise<boolean> {
    const documentId = Number(metricItem.DocumentId || 0);
    let encodedFileRef = '';
    let checkedOut = false;
 
    try {
      const digest = await this.getFreshDigest();
      const fileResp = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${KM_DATA_HUB_LIBRARY}')/items` +
        `?$filter=Id eq ${documentId}&$select=Id,FileRef&$top=1`,
        this.spHttpClientConfiguration
      );
      const fileData = await fileResp.json();
      const hubItem = fileData.d?.results?.[0] ?? fileData.value?.[0];
      if (!hubItem?.FileRef) {
        console.error('[SYNC] No FileRef for:', documentId);
        return false;
      }
 
      encodedFileRef = this.encodeFileRef(hubItem.FileRef);
      const checkoutResp = await this.retryRequest(() =>
        this.context.spHttpClient.post(
          `${this.siteUrl}/_api/web/GetFileByServerRelativeUrl('${encodedFileRef}')/checkout()`,
          this.spHttpClientConfiguration,
          { headers: { 'X-RequestDigest': digest }, body: '' }
        )
      );
      if (!checkoutResp.ok) {
        console.warn('[SYNC] Checkout failed or conflict, skipping:', documentId, checkoutResp.status);
        return false;
      }
      checkedOut = true;

      const metricsUpdateResp = await this.retryRequest(() =>
        this.context.spHttpClient.post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${KM_DATA_HUB_LIBRARY}')/items(${hubItem.Id})/ValidateUpdateListItem`,
          this.spHttpClientConfiguration,
          {
            headers: {
              'X-RequestDigest': digest,
              'Content-Type': 'application/json;odata=verbose',
              'Accept': 'application/json;odata=verbose',
              'OData-Version': ''
            },
            body: JSON.stringify({
              formValues: [
                { FieldName: 'Views', FieldValue: String(Number(metricItem.ViewCount || 0)) },
                { FieldName: 'Likes', FieldValue: String(Number(metricItem.LikeCount || 0)) },
                { FieldName: 'Downloads', FieldValue: String(Number(metricItem.DownloadCount || 0)) },
                { FieldName: 'Comments', FieldValue: String(Number(metricItem.CommentCount || 0)) },
                { FieldName: 'Share', FieldValue: String(Number(metricItem.ShareCount || 0)) },
                { FieldName: 'Bookmark', FieldValue: String(Number(metricItem.BookmarkCount || 0)) },
                { FieldName: 'Follow', FieldValue: String(Number(metricItem.FollowCount || 0)) }
              ],
              bNewDocumentUpdate: true,
              checkInComment: ''
            })
          }
        )
      );
      if (!metricsUpdateResp.ok) {
        console.error('[SYNC] Metrics ValidateUpdateListItem failed:', documentId, metricsUpdateResp.status);
        await this.undoCheckout(encodedFileRef);
        return false;
      }

      const metricsUpdateJson = await metricsUpdateResp.json();
      const fieldResults = metricsUpdateJson?.value ||
        metricsUpdateJson?.d?.ValidateUpdateListItem?.results ||
        metricsUpdateJson?.d?.results || [];
      const fieldErrors = fieldResults.filter((entry: any) => entry.HasException);
      if (fieldErrors.length > 0) {
        console.error(
          '[SYNC] Metrics field update failed:',
          documentId,
          fieldErrors.map((entry: any) => `${entry.FieldName}: ${entry.ErrorMessage}`).join('; ')
        );
        await this.undoCheckout(encodedFileRef);
        return false;
      }
 
      const checkinResp = await this.retryRequest(() =>
        this.context.spHttpClient.post(
          `${this.siteUrl}/_api/web/GetFileByServerRelativeUrl('${encodedFileRef}')/checkin(comment='Metrics sync',checkintype=2)`,
          this.spHttpClientConfiguration,
          { headers: { 'X-RequestDigest': digest }, body: '' }
        )
      );
      if (!checkinResp.ok) {
        console.error('[SYNC] Checkin failed:', documentId, checkinResp.status);
        await this.undoCheckout(encodedFileRef);
        return false;
      }
      checkedOut = false;
 
      await this.clearPendingSync(metricItem.Id);
      debugLog('[SYNC] Synced:', documentId);
      return true;
    } catch (err) {
      console.error('[SYNC] Error for:', documentId, err);
      if (checkedOut && encodedFileRef) {
        await this.undoCheckout(encodedFileRef);
      }
      return false;
    }
  }
 
  private async syncPendingMetrics(): Promise<void> {
    const inProgress = await this.getSyncConfig('SyncInProgress');
    if (inProgress === 'true') {
      this.showToast('Sync already in progress, please wait...');
      return;
    }
 
    const startTime = new Date();
    await this.setSyncConfig('SyncInProgress', 'true');
    await this.setSyncConfig('LastSyncStartedAt', startTime.toISOString());
    await this.setSyncConfig('LastSyncStatus', 'running');
    this.showToast('Sync started - fetching pending docs...');
 
    try {
      const pendingItems: any[] = [];
      let nextUrl: string | null =
        `${this.siteUrl}/_api/web/lists/getbytitle('${DOCUMENT_METRICS_LIST}')/items` +
        `?$filter=PendingSync eq 1` +
        `&$select=Id,DocumentId,ViewCount,LikeCount,DownloadCount,CommentCount,ShareCount,BookmarkCount,FollowCount` +
        `&$top=500`;
 
      while (nextUrl) {
        const resp = await this.context.spHttpClient.get(
          nextUrl,
          this.spHttpClientConfiguration
        );
        const data = await resp.json();
        const items = data.d?.results ?? data.value ?? [];
        pendingItems.push(...items);
        nextUrl = data.d?.__next ?? data['@odata.nextLink'] ?? null;
        if (nextUrl) {
          await new Promise(resolve => setTimeout(resolve, 300));
        }
      }

      debugLog('[SYNC] Pending docs:', pendingItems.length);
      if (pendingItems.length === 0) {
        this.showToast('Sync complete - no pending changes');
        await this.setSyncConfig('SyncInProgress', 'false');
        await this.setSyncConfig('LastSyncStatus', 'completed');
        await this.setSyncConfig('LastSyncedCount', '0');
        await this.setSyncConfig('LastSyncCompletedAt', new Date().toISOString());
        return;
      }
 
      this.showToast(`Syncing ${pendingItems.length} docs...`);
      const batchSize = 5;
      let synced = 0;
      let failed = 0;
 
      for (let i = 0; i < pendingItems.length; i += batchSize) {
        const batch = pendingItems.slice(i, i + batchSize);
        await Promise.all(batch.map(async (metricItem) => {
          const success = await this.syncOneToKMHub(metricItem);
          if (success) {
            synced++;
          } else {
            failed++;
          }
        }));
 
        if (i % 50 === 0 && i > 0) {
          this.showToast(`Syncing... ${Math.min(i + batchSize, pendingItems.length)}/${pendingItems.length}`);
        }
        await new Promise(resolve => setTimeout(resolve, 500));
      }
 
      const completedAt = new Date().toISOString();
      const duration = Math.round((new Date().getTime() - startTime.getTime()) / 1000);
      await this.setSyncConfig('LastSyncCompletedAt', completedAt);
      await this.setSyncConfig('LastSyncedAt', completedAt);
      await this.setSyncConfig('LastSyncedCount', String(synced));
      await this.setSyncConfig(
        'LastSyncStatus',
        failed > 0 ? `completed_with_${failed}_errors` : 'completed'
      );
      await this.setSyncConfig('SyncInProgress', 'false');
      this.showToast(
        `Sync complete - ${synced} docs synced in ${duration}s` +
        (failed > 0 ? ` (${failed} failed)` : '')
      );
      debugLog('[SYNC] Complete:', { synced, failed, duration });
    } catch (err) {
      console.error('[SYNC] Fatal:', err);
      await this.setSyncConfig('SyncInProgress', 'false');
      await this.setSyncConfig('LastSyncStatus', 'failed');
      this.showToast('Sync failed - check console for details');
    }
  }
}
 
 
