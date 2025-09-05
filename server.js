const https = require('https');
const fs = require('fs');
const path = require('path');
const express = require('express');

const app = express();

// Serve static files from dist directory
app.use(express.static(path.join(__dirname, 'dist')));

// Try to read the SSL certificates
let httpsOptions;
try {
  const certPath = path.join(require('os').homedir(), '.office-addin-dev-certs');
  httpsOptions = {
    key: fs.readFileSync(path.join(certPath, 'localhost.key')),
    cert: fs.readFileSync(path.join(certPath, 'localhost.crt'))
  };
} catch (error) {
  console.log('SSL certificates not found, using HTTP instead');
  console.log('For Office add-ins, you need HTTPS. Run: npx office-addin-dev-certs install');
}

const port = 3000;

if (httpsOptions) {
  https.createServer(httpsOptions, app).listen(port, () => {
    console.log(`HTTPS Server running at https://localhost:${port}`);
    console.log('Frontend ready for Office add-in development');
  });
} else {
  app.listen(port, () => {
    console.log(`HTTP Server running at http://localhost:${port}`);
    console.log('Warning: Office add-ins require HTTPS');
  });
}