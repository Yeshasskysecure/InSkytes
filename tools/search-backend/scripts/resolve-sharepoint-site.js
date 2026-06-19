const { loadEnv } = require('./lib/env');
const { getGraphToken, graphFetch } = require('./lib/graphClient');

const graphBaseUrl = 'https://graph.microsoft.com/v1.0';

const encodeSitePath = (sitePath) => {
  const normalized = String(sitePath || '').replace(/^\/+/, '');
  return normalized
    .split('/')
    .filter(Boolean)
    .map((segment) => encodeURIComponent(segment))
    .join('/');
};

const pickDriveForList = (drives, list) => {
  const webUrl = String(list.webUrl || '').replace(/\/Forms\/.*$/i, '').toLowerCase();
  return drives.find((drive) => String(drive.webUrl || '').toLowerCase() === webUrl)
    || drives.find((drive) => String(drive.name || '').toLowerCase() === String(list.displayName || '').toLowerCase())
    || null;
};

const run = async () => {
  const envPath = process.argv[2] || 'config/search.env';
  const hostname = process.argv[3] || 'skysecuretech.sharepoint.com';
  const sitePath = process.argv[4] || '/sites/InSkytes';
  const libraryName = process.argv[5] || 'KM Data Hub';
  const whosWhoName = process.argv[6] || "Who's Who";

  const { env, absolutePath } = loadEnv(envPath);
  const token = await getGraphToken(env);
  const site = await graphFetch(token, `${graphBaseUrl}/sites/${hostname}:/${encodeSitePath(sitePath)}?$select=id,displayName,webUrl`);
  const [lists, drives] = await Promise.all([
    graphFetch(token, `${graphBaseUrl}/sites/${encodeURIComponent(site.id)}/lists?$select=id,displayName,webUrl,list`),
    graphFetch(token, `${graphBaseUrl}/sites/${encodeURIComponent(site.id)}/drives?$select=id,name,webUrl`)
  ]);

  const findList = (name) => (lists.value || []).find((list) =>
    String(list.displayName || '').trim().toLowerCase() === String(name || '').trim().toLowerCase()
  ) || null;

  const library = findList(libraryName);
  const whosWho = findList(whosWhoName);
  const libraryDrive = library ? pickDriveForList(drives.value || [], library) : null;

  let librarySample;
  let whosWhoSample;

  if (library) {
    librarySample = await graphFetch(
      token,
      `${graphBaseUrl}/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(library.id)}/items?$expand=fields,driveItem&$top=3`
    );
  }

  if (whosWho) {
    whosWhoSample = await graphFetch(
      token,
      `${graphBaseUrl}/sites/${encodeURIComponent(site.id)}/lists/${encodeURIComponent(whosWho.id)}/items?$expand=fields&$top=3`
    );
  }

  console.log(`Using env: ${absolutePath}`);
  console.log(JSON.stringify({
    site,
    library: library && {
      id: library.id,
      displayName: library.displayName,
      webUrl: library.webUrl,
      itemCount: library.list && library.list.itemCount,
      driveId: libraryDrive && libraryDrive.id,
      driveName: libraryDrive && libraryDrive.name,
      sampleCount: (librarySample.value || []).length,
      sampleFiles: (librarySample.value || []).map((item) => ({
        id: item.id,
        fileName: item.fields && item.fields.FileLeafRef,
        status: item.fields && item.fields.Status,
        hasDriveItem: Boolean(item.driveItem)
      }))
    },
    whosWho: whosWho && {
      id: whosWho.id,
      displayName: whosWho.displayName,
      webUrl: whosWho.webUrl,
      itemCount: whosWho.list && whosWho.list.itemCount,
      sampleCount: (whosWhoSample.value || []).length,
      sampleTitles: (whosWhoSample.value || []).map((item) => item.fields && item.fields.Title)
    },
    availableDocumentLibraries: (lists.value || [])
      .filter((list) => list.list && list.list.template === 'documentLibrary')
      .map((list) => ({
        id: list.id,
        displayName: list.displayName,
        webUrl: list.webUrl,
        itemCount: list.list && list.list.itemCount
      })),
    availableDrives: (drives.value || []).map((drive) => ({
      id: drive.id,
      name: drive.name,
      webUrl: drive.webUrl
    }))
  }, null, 2));
};

run().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
