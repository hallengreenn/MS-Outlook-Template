# Troubleshooting Guide

Common issues and solutions for the Outlook Meeting Template Add-in.

## Table of Contents

- [Add-in doesn't appear in Outlook](#add-in-doesnt-appear-in-outlook)
- [Template doesn't insert automatically](#template-doesnt-insert-automatically)
- [Icons don't display](#icons-dont-display)
- [CORS errors](#cors-errors)
- [Deployment issues](#deployment-issues)
- [Getting diagnostic information](#getting-diagnostic-information)

---

## Add-in doesn't appear in Outlook

### Symptom
After installing the add-in, you don't see it in Outlook.

### Solutions

**1. Check if add-in is installed**
- Open Outlook
- Go to **File → Options → Add-ins**
- Look for "Meeting Template" in the list
- If not there, reinstall the add-in

**2. Check if add-in is disabled**
- In File → Options → Add-ins
- At the bottom, select **Manage: Disabled Items**
- Click **Go**
- If the add-in is listed, select it and click **Enable**

**3. Restart Outlook**
- Close Outlook completely (check Task Manager to ensure it's closed)
- Reopen Outlook
- Try creating a new meeting

**4. Clear Office cache**

Windows:
```
%LocalAppData%\Microsoft\Office\16.0\Wef\
```
Delete all files in this folder, then restart Outlook.

Mac:
```
~/Library/Containers/com.microsoft.Outlook/Data/Library/Caches/
```

**5. Check Office version**
- The add-in requires Office 2016 or newer with Mailbox API 1.10
- Update Office to the latest version

---

## Template doesn't insert automatically

### Symptom
The add-in is installed, but the template doesn't appear when creating a new meeting.

### Solutions

**1. Check Developer Console**
- Create a new meeting in Outlook
- Press **F12** to open Developer Tools
- Go to **Console** tab
- Look for errors (red text)
- Common errors:
  - `Failed to load resource: 404` - File not found
  - `CORS policy` - Server configuration issue
  - `onNewAppointmentOrganizer is not defined` - JavaScript error

**2. Verify event handler is registered**

In the Console, check if you see:
```
Office.js ready - Meeting Template Add-in loaded
Event handler registered: onNewAppointmentOrganizer
```

If not, the add-in didn't load correctly.

**3. Check manifest configuration**

Verify manifest.xml has the correct event configuration:
```xml
<LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onNewAppointmentOrganizer"/>
```

**4. Verify all URLs are correct**

In manifest.xml, check that all URLs point to your hosting:
- IconUrl
- Commands.Url
- Taskpane.Url
- WebViewRuntime.Url
- JSRuntime.Url

Open each URL in a browser to verify they load.

**5. Check template isn't already inserted**

The add-in checks if the meeting body already has content. If it does, it won't insert the template. Try:
- Creating a completely new meeting (don't reuse an existing one)
- Wait 2-3 seconds after opening the meeting window

**6. Clear browser cache**
- In Outlook, press **Ctrl+Shift+Delete**
- Or close Outlook and clear cache manually (see above)

---

## Icons don't display

### Symptom
The add-in works, but icons show as broken images or don't appear.

### Solutions

**1. Verify icon files exist**

Test each icon URL in a browser:
```
https://YOUR-DOMAIN.com/assets/icon-16.png
https://YOUR-DOMAIN.com/assets/icon-32.png
https://YOUR-DOMAIN.com/assets/icon-64.png
https://YOUR-DOMAIN.com/assets/icon-80.png
https://YOUR-DOMAIN.com/assets/icon-128.png
```

If any return 404, check:
- Files are uploaded to the server
- File names match exactly (case-sensitive)
- Folder structure is correct (`/assets/`)

**2. Check manifest.xml paths**

Verify icon URLs in manifest:
```xml
<IconUrl DefaultValue="https://YOUR-DOMAIN.com/assets/icon-64.png"/>
<HighResolutionIconUrl DefaultValue="https://YOUR-DOMAIN.com/assets/icon-128.png"/>

<!-- And in Resources: -->
<bt:Image id="Icon.16x16" DefaultValue="https://YOUR-DOMAIN.com/assets/icon-16.png"/>
<bt:Image id="Icon.32x32" DefaultValue="https://YOUR-DOMAIN.com/assets/icon-32.png"/>
<bt:Image id="Icon.80x80" DefaultValue="https://YOUR-DOMAIN.com/assets/icon-80.png"/>
```

**3. Check CORS headers**

Icons must be served with proper CORS headers. In `staticwebapp.config.json`:
```json
"globalHeaders": {
  "Access-Control-Allow-Origin": "*"
}
```

**4. Clear Office cache**

Sometimes Office caches old icon URLs. Clear cache (see above) and restart Outlook.

---

## CORS Errors

### Symptom
Console shows: `Access to fetch at '...' from origin '...' has been blocked by CORS policy`

### Solutions

**1. For Azure Static Web Apps**

Ensure `staticwebapp.config.json` exists in your deployment with:
```json
{
  "globalHeaders": {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Cache-Control": "public, max-age=3600"
  }
}
```

**2. For other hosting**

Add CORS headers to your server configuration:

**Apache (.htaccess):**
```apache
Header set Access-Control-Allow-Origin "*"
Header set Access-Control-Allow-Methods "GET, POST, OPTIONS"
```

**Nginx:**
```nginx
add_header Access-Control-Allow-Origin "*";
add_header Access-Control-Allow-Methods "GET, POST, OPTIONS";
```

**3. Verify CORS headers**

Test with curl:
```bash
curl -I https://YOUR-DOMAIN.com/commands.js
```

Look for:
```
Access-Control-Allow-Origin: *
```

---

## Deployment Issues

### Azure Static Web Apps

**Deployment failed**

1. Check GitHub Actions logs (if using GitHub integration):
   - Go to repository → Actions tab
   - Click on failed workflow
   - Review error messages

2. Common issues:
   - Wrong app location (should be `/` for root)
   - Wrong build preset (should be "Custom")
   - Missing files

**Files not found (404)**

1. Verify files are at root of deployment:
   ```
   /commands.html
   /commands.js
   /manifest.xml
   /assets/icon-16.png
   ```

2. Check Azure Portal:
   - Go to your Static Web App
   - Click **Browse**
   - Try accessing files directly

**Updates not appearing**

1. Wait 2-5 minutes for deployment to complete
2. Clear browser cache: Ctrl+Shift+R
3. Check GitHub Actions completed successfully
4. Verify version number incremented in manifest.xml

### GitHub Pages

**Add-in doesn't load**

1. Ensure HTTPS is enabled (required for Office Add-ins)
2. Check repository is public or Pages is enabled for private repos
3. Verify custom domain has SSL certificate

### General Hosting

**SSL Certificate issues**

Office Add-ins require valid HTTPS. Check:
- Certificate is not expired
- Certificate matches domain
- No mixed content (HTTP resources on HTTPS page)

---

## Getting Diagnostic Information

### Browser Console

1. Create new meeting in Outlook
2. Press **F12** to open Developer Tools
3. Go to **Console** tab
4. Look for messages starting with add-in name
5. Check for errors (red text)

### Network Tab

1. Press F12 → **Network** tab
2. Create new meeting
3. Look for requests to your domain
4. Check for:
   - Failed requests (red, 404, 500 errors)
   - CORS errors
   - Timing issues (very slow requests)

### Office Add-in Logging

Enable runtime logging:

**Windows:**
1. Create registry key:
   ```
   HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer
   ```
2. Add DWORD value:
   - Name: `EnableLogging`
   - Value: `1`

3. Logs appear in:
   ```
   %temp%\wef\
   ```

**Mac:**
1. Run in Terminal:
   ```bash
   defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true
   ```

2. View logs in Safari's Developer menu

### Manifest Validation

Validate your manifest.xml:
1. Use Office Add-in Validator:
   ```bash
   npx office-addin-manifest validate manifest.xml
   ```

2. Or upload to AppSource validation tool:
   https://appsource.microsoft.com/marketplace/apps

---

## Still Having Issues?

### Check the Basics

- [ ] Using HTTPS (required)
- [ ] All URLs in manifest.xml are correct
- [ ] Files uploaded to hosting
- [ ] Icons exist and are accessible
- [ ] Office is up to date
- [ ] Tested in multiple browsers/clients

### Get Help

1. Check the other documentation:
   - AZURE-DEPLOYMENT.md - Deployment guide
   - CONFIGURATION.md - Configuration options
   - DEPLOYMENT-GUIDE.md - General deployment

2. Check Office Add-ins documentation:
   - https://docs.microsoft.com/en-us/office/dev/add-ins/

3. Create an issue on GitHub:
   - Include error messages from Console
   - Include manifest.xml (with URLs redacted if needed)
   - Describe what you've tried

---

## Quick Checklist

Before asking for help, verify:

- [ ] Add-in is installed in Outlook (File → Options → Add-ins)
- [ ] Add-in is not disabled
- [ ] Outlook has been restarted
- [ ] Office cache has been cleared
- [ ] All URLs in manifest.xml are correct and accessible
- [ ] staticwebapp.config.json exists and has CORS headers
- [ ] Icons are uploaded and accessible
- [ ] Browser console shows no errors
- [ ] Using latest version of Office
- [ ] Tested creating a completely new meeting (not editing existing)

---

## Common Error Messages

### "Add-in Error: This add-in couldn't be started."

**Cause:** manifest.xml has errors or files can't be loaded.

**Solution:**
1. Validate manifest.xml
2. Check all URLs are accessible
3. Check browser console for specific errors

### "We're sorry, we can't load the add-in."

**Cause:** Files not found (404) or CORS issues.

**Solution:**
1. Verify all files are uploaded
2. Check CORS headers (see above)
3. Test URLs directly in browser

### "Office.js is not defined"

**Cause:** Office.js library didn't load.

**Solution:**
1. Check internet connection
2. Verify manifest.xml has correct Office.js reference
3. Check if using offline mode (not supported)

### Template inserts multiple times

**Cause:** Event handler being called multiple times.

**Solution:**
This is handled in the code by checking if body already has content. If still happening:
1. Check for duplicate event handler registrations
2. Verify only one version of add-in is installed
3. Clear Office cache

---

**Pro tip:** Most issues are caused by:
1. Incorrect URLs in manifest.xml
2. Missing CORS headers
3. Files not uploaded correctly
4. Office cache containing old versions

Start by checking these four things!
