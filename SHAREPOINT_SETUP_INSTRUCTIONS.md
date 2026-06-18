# SharePoint Setup Instructions for QuestionSection Tiles

## What You Need to Do:

### Step 1: Verify Your SharePoint List Columns

1. Go to your SharePoint list: `https://indegenedevelopment.sharepoint.com/sites/KnowledgeHub/KMArtifacts/Forms/AllItems.aspx`
2. Click on the **gear icon** (Settings) → **List settings**
3. Scroll down to **Columns** section
4. Note the exact column names, especially:
   - **Title/Name column** - What is the exact name? (Could be "Title", "Name", "TitleName", etc.)
   - **Abstract column** - What is the exact name? (Could be "Abstract", "Description", "Summary", etc.)

### Step 2: Check Browser Console

1. Open your SharePoint page with the web part
2. Press **F12** to open Developer Tools
3. Go to **Console** tab
4. Look for logs starting with `=== FETCHING DOCUMENTS ===`
5. Find the log that says **"First item sample:"** - this shows all available fields
6. **Copy and share that JSON object** - it will show the exact field names SharePoint is using

### Step 3: Verify Data Exists

1. Make sure you have at least 1 document in the KMArtifacts list
2. The document should have:
   - A file uploaded (PDF, DOCX, etc.)
   - A title/name filled in
   - An abstract/description (optional but recommended)

### Step 4: Common Field Name Mappings

SharePoint uses internal field names that might differ from display names:

- **Title field** → Usually just `Title` in API
- **Custom Title field** → Might be `TitleName` or `Title_x0020_Name` (with spaces encoded)
- **Abstract field** → Usually `Abstract` or might be `Abstract_x0020_Description`
- **File name** → `FileLeafRef` (this is automatic for document libraries)
- **File path** → `FileRef` or `ServerRelativeUrl`

### Step 5: If Tiles Still Don't Show Data

If after checking the console you see data but tiles are empty, share:
1. The "First item sample" JSON from console
2. Screenshot of your list columns
3. Any error messages from console

This will help me update the field mappings in the code to match your exact SharePoint setup.

## Quick Test:

1. Open browser console (F12)
2. Refresh the page
3. Look for: `=== FETCHING DOCUMENTS ===`
4. Check: `Response status: 200` (should be 200)
5. Check: `Items count: X` (should be > 0 if you have documents)
6. Copy the `First item sample:` JSON and share it

The code will automatically try multiple field name variations, but if your SharePoint uses different internal names, we'll need to update the code based on what the console shows.

