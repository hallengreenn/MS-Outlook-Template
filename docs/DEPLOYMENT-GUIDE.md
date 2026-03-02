# Deployment Guide

This guide will help you deploy the Outlook Meeting Template Add-in to production.

## Prerequisites

- A web server with HTTPS support
- Access to Microsoft 365 Admin Center (for centralized deployment)
- Or access to Intune (for device management deployment)

## Step 1: Host the Add-in

You need to host the add-in files on an HTTPS server. Here are some options:

### Option A: Azure Static Web Apps (Recommended)

1. Create an Azure Static Web App
2. Upload files from the `/deploy` folder (or root folder)
3. Note your URL (e.g., `https://your-app.azurestaticapps.net`)

### Option B: GitHub Pages

1. Enable GitHub Pages in your repository settings
2. Set source to main branch
3. Your URL will be `https://yourusername.github.io/outlook-meeting-template`

### Option C: Any HTTPS Server

You can use any web server that supports HTTPS:
- AWS S3 + CloudFront
- Netlify
- Vercel
- Your own web server

## Step 2: Update manifest.xml

Open `manifest.xml` and replace all instances of `https://YOUR-DOMAIN.com` with your actual domain:

```xml
<IconUrl DefaultValue="https://your-actual-domain.com/assets/icon-64.png"/>
<HighResolutionIconUrl DefaultValue="https://your-actual-domain.com/assets/icon-128.png"/>
<SupportUrl DefaultValue="https://your-actual-domain.com"/>
<AppDomain>https://your-actual-domain.com</AppDomain>
```

Also update in the Resources section:
```xml
<bt:Url id="Commands.Url" DefaultValue="https://your-actual-domain.com/commands.html"/>
<bt:Url id="Taskpane.Url" DefaultValue="https://your-actual-domain.com/taskpane.html"/>
<bt:Url id="WebViewRuntime.Url" DefaultValue="https://your-actual-domain.com/commands.html"/>
<bt:Url id="JSRuntime.Url" DefaultValue="https://your-actual-domain.com/commands.js"/>
```

Optionally update:
- `<ProviderName>` - Your organization name
- `<DisplayName>` - The name shown to users

## Step 3: Deploy to Users

### Option A: Microsoft 365 Admin Center (Centralized)

1. Go to: https://admin.microsoft.com
2. Navigate to **Settings → Integrated apps**
3. Click **Upload custom apps**
4. Upload your updated `manifest.xml`
5. Choose who gets the add-in:
   - **All users** - Deploy to entire organization
   - **Specific users/groups** - Deploy to pilot group first
   - **Just me** - Test deployment

6. Click **Deploy**

### Option B: Intune Policy

See `docs/INTUNE-POLICY-QUICK.md` for Intune deployment instructions.

### Option C: Manual Installation (Testing)

For testing or small deployments:
1. Send the manifest.xml to users
2. Users open Outlook Desktop
3. File → Get Add-ins → My Add-ins
4. Add a custom add-in → Add from file
5. Select manifest.xml

## Step 4: Test the Deployment

1. Open Outlook (Desktop, Web, or Mac)
2. Create a new meeting
3. Verify the template is inserted automatically
4. Check formatting is correct
5. Test in different Outlook clients if needed

## Step 5: Monitor and Support

- Monitor for errors in Microsoft 365 Admin Center
- Check add-in usage statistics
- Collect user feedback
- Update template as needed

## Updating the Add-in

To update the template or make changes:

1. Update `commands.js` with new template
2. Upload updated files to your web server
3. Increment version in manifest.xml (e.g., 1.0.0.0 → 1.1.0.0)
4. Update manifest in M365 Admin Center
5. Users will get the update automatically

## Troubleshooting Deployment

### Add-in doesn't appear for users

- Check that manifest.xml was deployed successfully in Admin Center
- Verify users are in the assigned group
- Check that HTTPS URLs are accessible
- Restart Outlook

### Template doesn't insert

- Check browser console (F12) for JavaScript errors
- Verify all files are uploaded correctly
- Check CORS settings on your server
- Verify staticwebapp.config.json is in place

### Icons don't load

- Verify icon files are in `/assets` folder
- Check URLs in manifest.xml are correct
- Test icon URLs directly in browser

## Security Considerations

- Always use HTTPS (required by Office Add-ins)
- Keep manifest.xml and source files in version control
- Test thoroughly before deploying to all users
- Consider a pilot group for initial deployment
- Monitor for security updates to Office.js

## Best Practices

1. **Test locally first** - Use `node server.js` for local testing
2. **Pilot deployment** - Deploy to a small group first
3. **Version control** - Track all changes in git
4. **Documentation** - Document any customizations
5. **User training** - Provide user guide (see `docs/BRUGER-GUIDE.md`)
6. **Feedback loop** - Collect and act on user feedback

## Support

For deployment issues:
- Check `docs/FIX-GUIDE.md` for troubleshooting
- Check Microsoft documentation: https://docs.microsoft.com/en-us/office/dev/add-ins/
- Create an issue in the GitHub repository
