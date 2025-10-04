// First, install the library:
// npm install qrcode

const QRCode = require('qrcode');

const url = "https://gokhalebandhu.com";

// Generate QR code and save as image
QRCode.toFile('gokhalebandhu_qr.png', url, {
  color: {
    dark: '#000000',  // QR code color
    light: '#FFFFFF'  // Background color
  },
  width: 300
}, function (err) {
  if (err) throw err;
  console.log('QR code saved as gokhalebandhu_qr.png');
});
