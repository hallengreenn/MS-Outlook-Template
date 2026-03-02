# Outlook Meeting Template Add-in

A professional Outlook add-in that automatically inserts structured meeting templates when creating new calendar appointments.

## Description

This open-source project makes it easy to standardize meeting documentation in your organization by automatically inserting a structured template into all new meeting appointments in Outlook.

**Default template includes:**
- Meeting purpose/objective
- Agenda items
- Roles and responsibilities
- Decisions and next steps
- Other business

## Features

- **Automatic activation** - Template is inserted automatically when creating a new meeting
- **Event-based** - Uses Office.js LaunchEvent API for fast response
- **Full HTML support** - Template supports formatting, lists, tables, etc.
- **Cross-platform** - Works in Outlook Desktop, Web, and Mac (when deployed to HTTPS)
- **Enterprise-ready** - Can be deployed centrally via Microsoft 365 Admin Center or Intune

## Quick Start

### Prerequisites

To deploy this add-in, you need:

- **HTTPS web hosting** - Required for Office Add-ins
  - Azure Static Web Apps (recommended - free tier available)
  - GitHub Pages
  - Netlify, Vercel, or any HTTPS server
- **Microsoft 365 Admin Center access** - For deploying to users
- **Outlook** (Desktop, Web, or Mac) - For testing
- **Code editor** (VS Code recommended) - For customization

### Deployment in 3 Steps

1. **Deploy to Azure Static Web Apps**
   - Follow the detailed guide: **docs/AZURE-DEPLOYMENT.md** ⭐
   - Upload all files from root directory to Azure
   - Or use GitHub Pages, Netlify, etc.

2. **Update manifest.xml**
   - Replace `https://YOUR-DOMAIN.com` with your actual hosting URL
   - Update `<ProviderName>` with your organization name

3. **Deploy to users**
   - Go to Microsoft 365 Admin Center
   - Navigate to **Settings → Integrated apps**
   - Upload your manifest.xml
   - Assign to users or groups

**See AZURE-DEPLOYMENT.md for detailed step-by-step instructions.**

### Local Testing (Optional)

For local development and testing before deployment:

1. **Install dependencies**
   ```bash
   npm install
   ```

2. **Start local server**
   ```bash
   node server.js
   ```

3. **Test in Outlook**
   - Use the localhost URLs in manifest.xml
   - Add to Outlook for testing: File → Get Add-ins → My Add-ins → Add from file

**Note:** Local testing is optional. You can deploy directly to Azure and test there.

## Project Structure

```
outlook-meeting-template/
│
├── README.md                    # This file
├── LICENSE                      # MIT License
│
├── manifest.xml                 # Office Add-in manifest (uses localhost for dev)
├── index.html                   # Landing page for add-in
├── commands.html                # Commands page (event handler)
├── commands.js                  # Event-based activation logic
├── taskpane.html                # Task pane UI
├── taskpane.js                  # Task pane logic
├── staticwebapp.config.json     # Azure Static Web App configuration
├── server.js                    # Local development server (optional)
├── package.json                 # NPM dependencies
│
├── assets/                      # Icons (PNG format)
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   ├── icon-80.png
│   └── icon-128.png
│
└── docs/                        # Documentation
    ├── AZURE-DEPLOYMENT.md      # Azure deployment guide
    ├── DEPLOYMENT-GUIDE.md      # General deployment guide
    ├── CONFIGURATION.md         # Configuration options
    └── TROUBLESHOOTING.md       # Troubleshooting guide
```

## Deployment

### Production Deployment

1. **Host the add-in**
   - Deploy all files from root directory to any HTTPS server
   - Azure Static Web Apps, GitHub Pages, Netlify, AWS S3, etc.

2. **Update manifest.xml**
   - Replace all `https://YOUR-DOMAIN.com` with your actual domain
   - Update `<ProviderName>` to your organization name
   - Update `<DisplayName>` if desired

3. **Deploy to users**
   - Via Microsoft 365 Admin Center: **Settings → Integrated apps → Upload custom app**
   - Assign to specific users, groups, or entire organization

### Deployment Guides

- **docs/AZURE-DEPLOYMENT.md** - Azure Static Web Apps deployment (recommended) ⭐
- **docs/DEPLOYMENT-GUIDE.md** - General deployment guide for any HTTPS server
- **docs/CONFIGURATION.md** - Detailed configuration options

## Customize the Template

The template can be customized in `commands.js`:

```javascript
const MEETING_TEMPLATE = `
<div style="font-family: 'Neue Haas Grotesk', Arial, sans-serif;">
<p><strong>Meeting Purpose</strong></p>
<p>Brief description of why we're meeting and what we want to achieve.</p>
<br>

<p><strong>Agenda Items</strong></p>
<p>List of topics to discuss (e.g., status updates or decisions)</p>
<br>

<p><strong>Roles and Responsibilities</strong></p>
<p>Who is the meeting leader and who takes notes (if relevant).</p>
<br>

<p><strong>Decisions and Next Steps</strong></p>
<p>Conclude by summarizing decisions and agreeing on follow-up.</p>
<br>

<p><strong>Other Business</strong></p>
<p>Time for questions or other items.</p>
</div>
`;
```

**After changes:**
1. Save the file
2. Restart local server (for testing)
3. Deploy updated files to production

## Technical Details

### Office.js API

The add-in uses the following Office.js APIs:
- **LaunchEvent API** - Event-based activation (OnNewAppointmentOrganizer)
- **Mailbox API** - Read/write to meeting body
- **Body API** - HTML formatting of body content

### Browser Compatibility

- Edge (Chromium)
- Chrome
- Safari (Mac)
- IE11 (legacy support via polyfills)

### Office Versions

- Microsoft 365 (Current Channel)
- Office 2021
- Office 2019
- Office 2016 (with updates)

## Troubleshooting

### Add-in doesn't appear in Outlook

1. Check that server is running (for local testing)
2. Verify manifest.xml is correctly installed
3. Restart Outlook
4. Check Event Viewer for errors

### Template doesn't insert automatically

1. Open Developer Tools (F12) in meeting window
2. Check Console for JavaScript errors
3. Verify LaunchEvent is registered correctly

See **docs/TROUBLESHOOTING.md** for comprehensive troubleshooting guide.

## Security

- Add-in uses HTTPS for production
- No sensitive data is stored
- Runs in sandboxed environment
- Follows Microsoft security best practices

## Contributing

Contributions are welcome! To contribute:
1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to branch
5. Create a Pull Request

## License

MIT License - See LICENSE file for details

## Version History

- **v1.2.0** - Font customization (Neue Haas Grotesk with fallback)
- **v1.1.0** - Event-based activation with LaunchEvent API
- **v1.0.0** - Initial release with manual activation

## Support

For questions or issues:
1. Check **docs/TROUBLESHOOTING.md** for common problems and solutions
2. Review deployment guides (AZURE-DEPLOYMENT.md, docs/DEPLOYMENT-GUIDE.md)
3. Check configuration options in docs/CONFIGURATION.md
4. Create an issue in the GitHub repository

## Credits

Developed as an open-source project to improve meeting culture and documentation across organizations.

## Related Projects

If you're interested in this project, you might also like:
- [Office Add-ins Samples](https://github.com/OfficeDev/Office-Add-in-samples)
- [Outlook Add-in Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/)
