# Azure Static Web Apps Deployment Guide

This guide walks you through deploying the Outlook Meeting Template Add-in to Azure Static Web Apps.

## Why Azure Static Web Apps?

Azure Static Web Apps is ideal for hosting Office Add-ins because:
- ✅ **Free tier available** - Perfect for internal tools
- ✅ **Automatic HTTPS** - Required for Office Add-ins
- ✅ **Global CDN** - Fast loading worldwide
- ✅ **Custom domains** - Use your own domain
- ✅ **Easy deployment** - GitHub integration or manual upload
- ✅ **No server management** - Fully managed service

## Prerequisites

- Azure account (create free at https://azure.microsoft.com/free/)
- Files from this repository
- (Optional) GitHub account for automatic deployment

## Deployment Options

### Option A: GitHub Actions (Recommended - Automatic)

This is the easiest way to deploy and get automatic updates when you push to GitHub.

#### Step 1: Fork/Clone Repository to GitHub

1. Fork this repository or push to your own GitHub repo
2. Make sure the repository is accessible to Azure

#### Step 2: Create Static Web App

1. Go to [Azure Portal](https://portal.azure.com)
2. Click **Create a resource**
3. Search for **Static Web App**
4. Click **Create**

#### Step 3: Configure Static Web App

**Basics tab:**
- **Subscription:** Your Azure subscription
- **Resource Group:** Create new or use existing
- **Name:** `outlook-meeting-template` (or your choice)
- **Plan type:** Free (for most use cases)
- **Region:** Choose closest to your users

**Deployment details:**
- **Source:** GitHub
- **Organization:** Your GitHub username
- **Repository:** Select your repository
- **Branch:** `main` (or `generic-version`)

**Build Details:**
- **Build Presets:** Custom
- **App location:** `/` (root of repository)
- **Api location:** (leave empty)
- **Output location:** (leave empty)

Click **Review + Create** → **Create**

#### Step 4: Wait for Deployment

- Azure will automatically deploy your app
- This takes 2-5 minutes
- Check the GitHub Actions tab in your repository to see progress

#### Step 5: Get Your URL

1. Go to your Static Web App in Azure Portal
2. Copy the URL (e.g., `https://nice-beach-04e97be03.1.azurestaticapps.net`)
3. Test by opening: `https://YOUR-URL.azurestaticapps.net/index.html`

---

### Option B: Azure CLI (Manual Upload)

Use this if you prefer command-line deployment or don't want GitHub integration.

#### Step 1: Install Azure CLI

Download and install from: https://aka.ms/installazurecli

#### Step 2: Login to Azure

```bash
az login
```

#### Step 3: Create Resource Group

```bash
az group create --name outlook-addin-rg --location westeurope
```

#### Step 4: Create Static Web App

```bash
az staticwebapp create \
  --name outlook-meeting-template \
  --resource-group outlook-addin-rg \
  --location westeurope \
  --sku Free
```

#### Step 5: Get Deployment Token

```bash
az staticwebapp secrets list \
  --name outlook-meeting-template \
  --resource-group outlook-addin-rg \
  --query "properties.apiKey" -o tsv
```

Copy this token - you'll need it for deployment.

#### Step 6: Deploy Files

Install the Static Web Apps CLI:

```bash
npm install -g @azure/static-web-apps-cli
```

Deploy your files:

```bash
cd /path/to/outlook-meeting-template
swa deploy ./  --deployment-token YOUR_DEPLOYMENT_TOKEN
```

---

### Option C: Azure Portal Upload (Simplest)

The easiest way for one-time deployment without GitHub or CLI.

#### Step 1: Prepare Files

1. Create a ZIP file containing:
   - All files from root directory (commands.js, manifest.xml, etc.)
   - assets/ folder
   - docs/ folder (optional)
   - staticwebapp.config.json

2. Make sure these files are at the root of the ZIP (not in a subfolder)

#### Step 2: Create Static Web App

1. Go to [Azure Portal](https://portal.azure.com)
2. Create a new **Static Web App**
3. Choose **Other** for deployment source
4. Complete the creation

#### Step 3: Upload Files

**Option 3a: Using Azure Portal:**
1. Go to your Static Web App
2. Click **Browse** to open the site
3. If you see a placeholder page, continue to next steps

**Option 3b: Using SWA CLI:**
```bash
npm install -g @azure/static-web-apps-cli
swa deploy ./ --env production
```

---

## Post-Deployment Configuration

### Step 1: Update manifest.xml

Now that you have your Azure URL, update the manifest:

1. Open `manifest.xml`
2. Replace all `https://YOUR-DOMAIN.com` with your actual Azure URL
3. Example: `https://nice-beach-04e97be03.1.azurestaticapps.net`

**Files to update:**
- Root `manifest.xml`
- `deploy/manifest.xml` (if using deploy folder)

**What to replace:**
```xml
<!-- Before -->
<IconUrl DefaultValue="https://YOUR-DOMAIN.com/assets/icon-64.png"/>

<!-- After -->
<IconUrl DefaultValue="https://nice-beach-04e97be03.1.azurestaticapps.net/assets/icon-64.png"/>
```

Do this for all URLs in the manifest:
- IconUrl
- HighResolutionIconUrl
- SupportUrl
- AppDomain
- All URLs in Resources section (Commands.Url, Taskpane.Url, etc.)

### Step 2: Redeploy with Updated Manifest

If using GitHub Actions:
```bash
git add manifest.xml
git commit -m "Update manifest with Azure URL"
git push
```

If using manual upload:
- Re-upload the ZIP with updated manifest.xml
- Or use SWA CLI to redeploy

### Step 3: Test the Deployment

1. Open: `https://YOUR-URL.azurestaticapps.net/index.html`
2. Verify all files load correctly
3. Check: `https://YOUR-URL.azurestaticapps.net/manifest.xml`
4. Verify all icon URLs: `https://YOUR-URL.azurestaticapps.net/assets/icon-64.png`

### Step 4: Test in Outlook

**Manual Installation (for testing):**
1. Download the updated manifest.xml from your Azure URL
2. Open Outlook Desktop
3. File → Get Add-ins → My Add-ins
4. Add a custom add-in → Add from URL
5. Enter: `https://YOUR-URL.azurestaticapps.net/manifest.xml`
6. Or use "Add from file" with downloaded manifest.xml

**Test the add-in:**
1. Create a new meeting in Outlook
2. Verify template is inserted automatically
3. Check formatting is correct

---

## Custom Domain (Optional)

Want to use your own domain instead of `*.azurestaticapps.net`?

### Step 1: Configure Custom Domain in Azure

1. Go to your Static Web App in Azure Portal
2. Click **Custom domains** in left menu
3. Click **+ Add**
4. Enter your domain: `addins.yourcompany.com`
5. Choose DNS validation method (CNAME or TXT)

### Step 2: Update DNS

Add the required DNS records at your domain registrar:

**CNAME Record:**
```
Name: addins
Type: CNAME
Value: nice-beach-04e97be03.1.azurestaticapps.net
TTL: 3600
```

**TXT Record (for validation):**
```
Name: _dnsauth.addins
Type: TXT
Value: [value from Azure Portal]
TTL: 3600
```

### Step 3: Wait for Validation

- DNS propagation takes 5-60 minutes
- Azure will automatically validate and issue SSL certificate
- Your add-in will be available at your custom domain

### Step 4: Update Manifest

Update all URLs in manifest.xml to use your custom domain:
```xml
<IconUrl DefaultValue="https://addins.yourcompany.com/assets/icon-64.png"/>
```

---

## Updating the Add-in

To update your add-in after initial deployment:

### If Using GitHub Actions:

1. Make changes to template in `commands.js`
2. Update version in `manifest.xml`: `1.2.0.0` → `1.3.0.0`
3. Commit and push:
   ```bash
   git add commands.js manifest.xml
   git commit -m "Update meeting template"
   git push
   ```
4. GitHub Actions will automatically deploy
5. Changes are live in 2-5 minutes

### If Using Manual Upload:

1. Make changes locally
2. Update version in manifest.xml
3. Re-upload files using Azure CLI or Portal
4. Users will see updates automatically

---

## Troubleshooting

### Deployment Failed

**Check GitHub Actions logs:**
1. Go to your repository on GitHub
2. Click **Actions** tab
3. Click on the failed workflow
4. Check error messages

**Common issues:**
- Missing staticwebapp.config.json
- Incorrect app location (should be `/`)
- Build preset not set to "Custom"

### Files Not Loading (404 Errors)

1. Check files are at root of repository, not in subfolder
2. Verify staticwebapp.config.json is present
3. Check Azure Portal → Static Web App → Configuration
4. Test individual file URLs in browser

### Icons Not Displaying

1. Verify icon files exist: `https://YOUR-URL/assets/icon-64.png`
2. Check file names are lowercase (icon-64.png, not Icon-64.png)
3. Verify paths in manifest.xml match actual file structure
4. Clear browser cache and try again

### Template Not Inserting

1. Open browser Developer Tools (F12) in Outlook
2. Check Console for JavaScript errors
3. Verify commands.js loaded correctly
4. Check Network tab for failed requests

### CORS Errors

If you see CORS errors in console:

1. Check staticwebapp.config.json has correct headers:
   ```json
   "globalHeaders": {
     "Access-Control-Allow-Origin": "*"
   }
   ```
2. Redeploy after making changes

---

## Cost Optimization

### Free Tier Limits

Azure Static Web Apps Free tier includes:
- ✅ 100 GB bandwidth/month
- ✅ Unlimited hosting
- ✅ Custom domains
- ✅ Automatic SSL certificates

For a small organization (100-500 users), Free tier is sufficient.

### Monitoring Usage

1. Go to Azure Portal → your Static Web App
2. Click **Metrics**
3. Monitor:
   - Data Out (bandwidth)
   - Requests
   - Errors

### Upgrade if Needed

If you exceed Free tier limits, upgrade to Standard tier:
- Standard tier: ~$9/month
- Includes more bandwidth and advanced features

---

## Security Best Practices

1. **Enable HTTPS only** (automatic with Azure Static Web Apps)
2. **Use custom domain** for professional appearance
3. **Monitor access logs** in Azure Portal
4. **Keep manifest.xml in version control**
5. **Test updates in staging before production**
6. **Use secrets for sensitive configuration** (if needed later)

---

## Production Checklist

Before deploying to all users:

- [ ] Tested add-in locally with `node server.js`
- [ ] Updated all URLs in manifest.xml to Azure URL
- [ ] Verified icons load correctly
- [ ] Tested in Outlook Desktop
- [ ] Tested in Outlook Web
- [ ] (Optional) Tested in Outlook Mac
- [ ] Customized template for your organization
- [ ] Updated ProviderName in manifest.xml
- [ ] Incremented version number
- [ ] Documented any customizations
- [ ] Created user guide (see docs/BRUGER-GUIDE.md)
- [ ] Prepared IT support documentation

---

## Next Steps

After successful Azure deployment:

1. **Deploy to users** via Microsoft 365 Admin Center
   - See `DEPLOYMENT-GUIDE.md` for instructions

2. **Or deploy via Intune** for managed devices
   - See `docs/INTUNE-POLICY-QUICK.md`

3. **Monitor adoption** and collect feedback

4. **Iterate and improve** template based on user needs

---

## Support

For Azure-specific issues:
- [Azure Static Web Apps Documentation](https://docs.microsoft.com/en-us/azure/static-web-apps/)
- [Azure Support](https://azure.microsoft.com/support/)

For add-in issues:
- See `DEPLOYMENT-GUIDE.md`
- See `docs/FIX-GUIDE.md`
- Create an issue in GitHub repository

---

## Frequently Asked Questions

**Q: Can I use Azure Static Web Apps with a free Azure account?**
A: Yes! The Free tier is perfect for internal add-ins.

**Q: How long does deployment take?**
A: Initial deployment: 5-10 minutes. Updates: 2-5 minutes.

**Q: Can I use a different Azure region?**
A: Yes, choose the region closest to your users for best performance.

**Q: Do I need a custom domain?**
A: No, the `*.azurestaticapps.net` domain works fine. Custom domains are optional.

**Q: How do I roll back a deployment?**
A: If using GitHub Actions, revert your git commit and push. Azure will redeploy the previous version.

**Q: Can I have staging and production environments?**
A: Yes! Create two Static Web Apps (staging and prod) and deploy different branches to each.

**Q: What if I exceed Free tier limits?**
A: Upgrade to Standard tier ($9/month) or optimize asset loading.
