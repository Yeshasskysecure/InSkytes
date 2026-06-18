# How SharePoint Connection Works for Tiles

## Automatic Site Detection

The code **automatically** connects to the SharePoint site where the web part is deployed. No manual URL configuration is needed!

### How It Works:

1. **SPFx Context Provides Site URL Automatically**
   ```typescript
   const webUrl = props.context.pageContext.web.absoluteUrl;
   ```
   - If your web part is on: `https://indegenedevelopment.sharepoint.com/sites/KnowledgeHub`
   - Then `webUrl` = `https://indegenedevelopment.sharepoint.com/sites/KnowledgeHub`
   - This is **automatically detected** by SharePoint Framework

2. **API Endpoint Construction**
   ```typescript
   const apiUrl = `${webUrl}/_api/web/lists/getbytitle('KMArtifacts')/items?...`;
   ```
   - Results in: `https://indegenedevelopment.sharepoint.com/sites/KnowledgeHub/_api/web/lists/getbytitle('KMArtifacts')/items?...`
   - SharePoint REST API uses this to find the library on the current site

## What You Need:

### ✅ Required:
1. **Web part deployed on the same site** where KMArtifacts library exists
   - Site: `https://indegenedevelopment.sharepoint.com/sites/KnowledgeHub`
   - Library: `KMArtifacts` (must exist on this site)

2. **Library name must be exactly**: `KMArtifacts`
   - Case-sensitive
   - Must match the display name in SharePoint

3. **User permissions**: The logged-in user must have **Read** access to the KMArtifacts library

### ❌ NOT Needed:
- ❌ Manual URL configuration
- ❌ Hardcoded site URLs
- ❌ Additional configuration files
- ❌ The full library URL (just the library name is enough)

## How to Verify It's Working:

1. **Check Browser Console (F12)**
   - Look for: `Current Site URL (auto-detected): https://indegenedevelopment.sharepoint.com/sites/KnowledgeHub`
   - Look for: `Expected Library URL: https://indegenedevelopment.sharepoint.com/sites/KnowledgeHub/KMArtifacts`
   - Look for: `Full API Endpoint: ...`

2. **Verify Site Match**
   - The "Current Site URL" should match your site: `https://indegenedevelopment.sharepoint.com/sites/KnowledgeHub`
   - If it doesn't match, the web part is on a different site

3. **Check Response**
   - Look for: `Response status: 200` (success)
   - Look for: `Items count: X` (should be > 0 if documents exist)

## Troubleshooting:

### Issue: "No documents found" but library exists
**Possible causes:**
1. Web part is on a different site than the library
2. Library name doesn't match exactly (case-sensitive)
3. User doesn't have read permissions
4. Library is empty

**Solution:** Check console logs to see the exact site URL being used

### Issue: 404 Error
**Possible causes:**
1. Library doesn't exist on the current site
2. Library name is misspelled
3. Site URL is incorrect

**Solution:** Verify the library exists on the site shown in "Current Site URL"

### Issue: 403 Forbidden
**Possible causes:**
1. User doesn't have permissions to read the library
2. Library has restricted access

**Solution:** Check user permissions in SharePoint

## Connection Flow Diagram:

```
User opens page with web part
    ↓
SPFx automatically detects current site URL
    ↓
webUrl = "https://indegenedevelopment.sharepoint.com/sites/KnowledgeHub"
    ↓
API call: webUrl + "/_api/web/lists/getbytitle('KMArtifacts')/items"
    ↓
SharePoint REST API finds library on current site
    ↓
Returns latest 3 documents
    ↓
Tiles display the data
```

## Summary:

**The link you provided (`https://indegenedevelopment.sharepoint.com/sites/KnowledgeHub/KMArtifacts/Forms/AllItems.aspx`) is enough!**

The code automatically:
- ✅ Detects the site: `https://indegenedevelopment.sharepoint.com/sites/KnowledgeHub`
- ✅ Finds the library: `KMArtifacts`
- ✅ Fetches the latest 3 documents
- ✅ Displays them in tiles

**No additional configuration needed** - just make sure:
1. Web part is deployed on the same site
2. Library name is exactly `KMArtifacts`
3. User has read permissions

