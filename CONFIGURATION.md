# Configuration Guide

This guide explains how to configure the Outlook Meeting Template Add-in for your organization.

## Manifest Configuration

The `manifest.xml` file contains all the configuration for your add-in. Here are the key settings to customize:

### Basic Information

```xml
<Id>d6ff5248-95b4-4400-8f5d-4ce60c4ed0be</Id>
```
- **What:** Unique identifier for your add-in
- **Change:** Keep the default, or generate a new GUID if forking
- **How:** Use online GUID generator or PowerShell: `[guid]::NewGuid()`

```xml
<Version>1.2.0.0</Version>
```
- **What:** Add-in version number
- **Change:** Increment when you make updates (e.g., 1.2.0.0 → 1.3.0.0)
- **Why:** Helps track deployments and updates

```xml
<ProviderName>Your Organization</ProviderName>
```
- **What:** Your organization name
- **Change:** Replace "Your Organization" with your company name
- **Example:** `<ProviderName>Contoso Inc</ProviderName>`

```xml
<DisplayName DefaultValue="Meeting Template"/>
```
- **What:** Name shown to users in Outlook
- **Change:** Customize if desired
- **Example:** `<DisplayName DefaultValue="Contoso Meeting Helper"/>`

```xml
<Description DefaultValue="Automatically inserts meeting template when creating appointments"/>
```
- **What:** Description shown in add-in management
- **Change:** Customize for your organization

### URLs and Domains

All URLs must be updated to point to your hosting location:

```xml
<IconUrl DefaultValue="https://YOUR-DOMAIN.com/assets/icon-64.png"/>
<HighResolutionIconUrl DefaultValue="https://YOUR-DOMAIN.com/assets/icon-128.png"/>
<SupportUrl DefaultValue="https://YOUR-DOMAIN.com"/>
```

Replace `https://YOUR-DOMAIN.com` with your actual domain:
- Azure Static Web App: `https://your-app.azurestaticapps.net`
- GitHub Pages: `https://yourusername.github.io/outlook-meeting-template`
- Custom domain: `https://addins.yourcompany.com`

```xml
<AppDomain>https://YOUR-DOMAIN.com</AppDomain>
```
- **What:** Authorized domain for the add-in
- **Must match:** The domain where you host the add-in files

### Resources (URLs)

In the `<Resources>` section at the bottom, update all URLs:

```xml
<bt:Url id="Commands.Url" DefaultValue="https://YOUR-DOMAIN.com/commands.html"/>
<bt:Url id="Taskpane.Url" DefaultValue="https://YOUR-DOMAIN.com/taskpane.html"/>
<bt:Url id="WebViewRuntime.Url" DefaultValue="https://YOUR-DOMAIN.com/commands.html"/>
<bt:Url id="JSRuntime.Url" DefaultValue="https://YOUR-DOMAIN.com/commands.js"/>
```

### Icons

Update icon URLs:

```xml
<bt:Image id="Icon.16x16" DefaultValue="https://YOUR-DOMAIN.com/assets/icon-16.png"/>
<bt:Image id="Icon.32x32" DefaultValue="https://YOUR-DOMAIN.com/assets/icon-32.png"/>
<bt:Image id="Icon.80x80" DefaultValue="https://YOUR-DOMAIN.com/assets/icon-80.png"/>
```

## Template Configuration

The meeting template is defined in `commands.js`:

### Basic Template Structure

```javascript
const MEETING_TEMPLATE = `
<div style="font-family: 'Neue Haas Grotesk', Arial, sans-serif;">
<p><strong>Section Title</strong></p>
<p>Section content</p>
<br>
</div>
`;
```

### Customization Options

#### Change Font

```javascript
// Default
font-family: 'Neue Haas Grotesk', Arial, sans-serif;

// Use your corporate font
font-family: 'Your Corporate Font', Arial, sans-serif;

// Simple sans-serif
font-family: Arial, sans-serif;
```

#### Add Sections

```javascript
const MEETING_TEMPLATE = `
<div style="font-family: 'Neue Haas Grotesk', Arial, sans-serif;">

<p><strong>Custom Section</strong></p>
<p>Your custom content</p>
<br>

<p><strong>Another Section</strong></p>
<ul>
  <li>Bullet point 1</li>
  <li>Bullet point 2</li>
</ul>
<br>

</div>
`;
```

#### Add Colors

```javascript
<p style="color: #0078d4;"><strong>Blue Heading</strong></p>
<p style="background-color: #f0f0f0; padding: 10px;">Highlighted text</p>
```

#### Add Tables

```javascript
<table border="1" cellpadding="5" cellspacing="0">
  <tr>
    <th>Column 1</th>
    <th>Column 2</th>
  </tr>
  <tr>
    <td>Data 1</td>
    <td>Data 2</td>
  </tr>
</table>
```

## Server Configuration

### staticwebapp.config.json

This file configures the web server (for Azure Static Web Apps):

```json
{
  "platformErrorOverrides": [
    {
      "errorType": "NotFound",
      "serve": "/index.html",
      "statusCode": 404
    }
  ],
  "mimeTypes": {
    ".json": "application/json",
    ".xml": "application/xml"
  },
  "globalHeaders": {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Cache-Control": "public, max-age=3600"
  }
}
```

**Key settings:**
- `Access-Control-Allow-Origin: *` - Allows Outlook to load the add-in
- `Cache-Control` - Controls caching (adjust as needed)

## Locale Configuration

To support multiple languages, you can add locale-specific versions:

```xml
<DisplayName DefaultValue="Meeting Template">
  <Override Locale="da-DK" Value="Møde Skabelon"/>
  <Override Locale="de-DE" Value="Besprechungsvorlage"/>
</DisplayName>
```

And create locale-specific templates:

```javascript
const LOCALE = Office.context.displayLanguage || 'en-US';

const TEMPLATES = {
  'en-US': `<p><strong>Meeting Purpose</strong></p>...`,
  'da-DK': `<p><strong>Formål med mødet</strong></p>...`,
  'de-DE': `<p><strong>Besprechungszweck</strong></p>...`
};

const MEETING_TEMPLATE = TEMPLATES[LOCALE] || TEMPLATES['en-US'];
```

## Advanced Configuration

### Event Configuration

The add-in uses event-based activation. This is configured in the manifest:

```xml
<LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onNewAppointmentOrganizer"/>
```

Available events:
- `OnNewAppointmentOrganizer` - When user creates new meeting (current)
- `OnAppointmentSend` - When user sends meeting invite
- `OnAppointmentAttachmentsChanged` - When attachments change

### Runtime Configuration

```xml
<Runtime resid="WebViewRuntime.Url">
  <Override type="javascript" resid="JSRuntime.Url"/>
</Runtime>
```

This tells Outlook where to find the JavaScript runtime.

## Testing Configuration

For local development, the manifest can use `localhost`:

```xml
<IconUrl DefaultValue="http://localhost:3000/assets/icon-64.png"/>
```

For production, always use HTTPS:

```xml
<IconUrl DefaultValue="https://your-domain.com/assets/icon-64.png"/>
```

## Validation

Before deploying, validate your configuration:

1. **Manifest validation:**
   - Use Office Add-in Validator: https://github.com/OfficeDev/office-addin-manifest
   - Or upload to Outlook to test

2. **URL validation:**
   - Test all URLs in browser
   - Verify HTTPS certificates are valid
   - Check CORS headers

3. **Template validation:**
   - Test in Outlook Desktop, Web, and Mac
   - Verify formatting renders correctly
   - Test with different themes (light/dark)

## Troubleshooting Configuration

### Icons don't load
- Check URLs are correct
- Verify files exist at those URLs
- Check file extensions (.png, not .PNG)
- Test URLs directly in browser

### Template doesn't appear
- Check JavaScript console (F12) for errors
- Verify commands.js is loaded
- Check event handler is registered

### CORS errors
- Check staticwebapp.config.json has correct CORS headers
- Verify server allows cross-origin requests
- Test with browser DevTools Network tab

## Best Practices

1. **Version control:** Keep manifest.xml in git with placeholder URLs
2. **Environment-specific:** Use different manifests for dev/staging/prod
3. **Documentation:** Document any customizations you make
4. **Testing:** Always test locally before deploying
5. **Backup:** Keep backup of working configuration

## Support

For configuration help:
- See `docs/DEPLOYMENT-GUIDE.md`
- See `docs/FIX-GUIDE.md`
- Create an issue in GitHub repository
