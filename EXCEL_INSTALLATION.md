# Excel Add-in Installation Guide

## Method 1: Local Manifest Installation (Recommended)

### Step 1: Download the Manifest
1. Go to your GitHub repository: `https://github.com/Everett022/COR-AUTO`
2. Download the `manifest.xml` file from the root directory
3. Save it to your desktop or a known location

### Step 2: Create the WEF Folder
1. Open Terminal (Applications → Utilities → Terminal)
2. Run this command to create the WEF folder:
   ```bash
   mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
   ```

### Step 3: Copy the Manifest
1. Copy the `manifest.xml` file to the WEF folder:
   ```bash
   cp ~/Desktop/manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
   ```
   (Replace `~/Desktop/` with the actual path where you saved the manifest)

### Step 4: Restart Excel
1. Close Excel completely (Cmd+Q)
2. Reopen Excel
3. Go to the **Home** tab
4. Look for a new group called **"Automation Tools"**
5. You should see a **"COR-AUTO"** button

### Step 5: Test the Add-in
1. Click the **"COR-AUTO"** button in the Automation Tools group
2. A taskpane should open on the right side of Excel
3. The add-in should now be functional!

## Method 2: Shared Network Installation (For Teams)

If you want to share this with colleagues:

1. **Host the manifest on a network share** or web server
2. **Have each user add it via Excel**:
   - Excel → Insert → My Add-ins → Upload My Add-in
   - Browse to the manifest.xml file
   - Click "Upload"

## Troubleshooting

### If the add-in doesn't appear:
1. **Check the WEF folder exists**:
   ```bash
   ls ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
   ```

2. **Verify the manifest is there**:
   ```bash
   ls ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/manifest.xml
   ```

3. **Check Excel's add-in list**:
   - Excel → Insert → My Add-ins
   - Look for "COR-AUTO" in the list

4. **Restart Excel completely**:
   - Force quit Excel (Cmd+Option+Esc)
   - Reopen Excel

### If you get errors:
1. **Check the manifest URLs**: Make sure they point to `https://everett022.github.io/COR-AUTO/`
2. **Verify GitHub Pages is working**: Visit `https://everett022.github.io/COR-AUTO/` in your browser
3. **Check Excel's error logs**: Excel → Help → Show Diagnostics

## File Locations

- **Manifest location**: `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/manifest.xml`
- **GitHub Pages URL**: `https://everett022.github.io/COR-AUTO/`
- **Repository**: `https://github.com/Everett022/COR-AUTO`

## Support

If you continue to have issues:
1. Check that GitHub Pages is deployed successfully
2. Verify all URLs in the manifest are accessible
3. Make sure Excel has internet access to load the add-in content 