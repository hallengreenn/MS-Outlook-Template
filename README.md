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

- Node.js installed (https://nodejs.org)
- Outlook Desktop (Microsoft 365 or Office 2016+)
- Code editor (VS Code recommended)

### Local Development in 3 Steps

1. **Install dependencies**
   ```bash
   npm install
   ```

2. **Start local server**
   ```bash
   node server.js
   ```
   Server will run on http://localhost:3000

3. **Add add-in to Outlook**
   - Open Outlook Desktop
   - Go to **File** → **Get Add-ins** → **My Add-ins**
   - Select **Add a custom add-in** → **Add from file...**
   - Select `manifest.xml` from this project
   - Click **Install**

4. **Test it!**
   - Go to Calendar in Outlook
   - Click **New Meeting**
   - Template is inserted automatically

**See `docs/QUICKSTART.md` for detailed guide.**

## Project Structure

```
outlook-meeting-template/
│
├── manifest.xml                 # Office Add-in manifest (uses localhost for dev)
├── index.html                   # Landing page for add-in
├── commands.html                # Commands page (event handler)
├── commands.js                  # Event-based activation logic
├── taskpane.html                # Task pane UI
├── taskpane.js                  # Task pane logic
├── staticwebapp.config.json     # Azure Static Web App configuration
├── server.js                    # Local development server
├── package.json                 # NPM dependencies
│
├── assets/                      # Icons (PNG format)
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   ├── icon-80.png
│   └── icon-128.png
│
├── deploy/                      # Ready for production deployment
│   ├── manifest.xml             # Manifest with YOUR-DOMAIN.com placeholders
│   ├── index.html
│   ├── commands.html
│   ├── commands.js
│   ├── taskpane.html
│   ├── taskpane.js
│   ├── staticwebapp.config.json
│   └── assets/                  # Icons
│
└── docs/                        # Documentation
    ├── QUICKSTART.md            # Quick start guide
    ├── DEPLOYMENT.md            # Full deployment guide
    └── ...
```

## Deployment

### Production Deployment

1. **Host the add-in**
   - Deploy files to any HTTPS server (Azure Static Web Apps, GitHub Pages, AWS S3, etc.)
   - Or use the `/deploy` folder for a clean deployment package

2. **Update manifest.xml**
   - Replace all `https://YOUR-DOMAIN.com` with your actual domain
   - Update `<ProviderName>` to your organization name
   - Update `<DisplayName>` if desired

3. **Deploy to users**
   - Via Microsoft 365 Admin Center: **Settings → Integrated apps → Upload custom app**
   - Via Intune: See `docs/INTUNE-POLICY-QUICK.md`

### Deployment Guides

- **AZURE-DEPLOYMENT.md** - Azure Static Web Apps deployment (recommended) ⭐
- **DEPLOYMENT-GUIDE.md** - General deployment guide for any HTTPS server
- **CONFIGURATION.md** - Detailed configuration options
- **docs/INTUNE-POLICY-QUICK.md** - Deploy via Intune to all users

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

See `docs/FIX-GUIDE.md` for more troubleshooting tips.

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
1. Check documentation in `/docs` folder
2. See troubleshooting guide: `docs/FIX-GUIDE.md`
3. Create an issue in the GitHub repository

## Credits

Developed as an open-source project to improve meeting culture and documentation across organizations.

## Related Projects

If you're interested in this project, you might also like:
- [Office Add-ins Samples](https://github.com/OfficeDev/Office-Add-in-samples)
- [Outlook Add-in Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/)
