const express = require('express');
const path = require('path');
const https = require('https');
const fs = require('fs');

const app = express();
const PORT = 3000;

// Serve static files
app.use(express.static(__dirname));

// Serve manifest.xml
app.get('/manifest.xml', (req, res) => {
  res.sendFile(path.join(__dirname, 'manifest.xml'));
});

// For development - create self-signed cert if needed
// I produktion skal du bruge et rigtigt SSL certifikat

app.listen(PORT, () => {
  console.log(`
╔════════════════════════════════════════════════════════════╗
║  Møde Skabelon Add-in Server                               ║
╚════════════════════════════════════════════════════════════╝

✅ Server kører på: http://localhost:${PORT}

📋 Næste skridt:
1. Kør: npm install
2. Kør: npm start
3. Åbn Outlook desktop
4. Gå til File > Get Add-ins > My Add-ins > Add a custom add-in > Add from file
5. Vælg manifest.xml fra denne mappe

💡 Til produktion deployment via Intune:
   - Host filerne på en HTTPS server (påkrævet for produktion)
   - Upload manifest via Microsoft 365 Admin Center
   - Deploy via Intune til dine brugere

  `);
});
