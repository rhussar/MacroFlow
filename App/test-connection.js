/**
 * Smoke Test: Excel COM Connection
 * Run with: node test-connection.js
 *
 * NOTE: This must be run with the same Node version that winax was compiled for.
 * If you rebuilt winax for Electron, this won't work from regular Node.
 * In that case, test from within Electron's main process.
 */

console.log('='.repeat(60));
console.log('Excel COM Connection Smoke Test');
console.log('='.repeat(60));
console.log('');

// Step 1: Check winax is installed
let winax;
try {
  winax = require('winax');
  console.log('[OK] winax loaded successfully');
  console.log('     Exports:', Object.keys(winax).join(', '));
} catch (err) {
  console.log('[FAIL] Could not load winax:', err.message);
  console.log('');
  console.log('If you see "NODE_MODULE_VERSION" error:');
  console.log('  - winax was compiled for a different Node version');
  console.log('  - Run: npx electron-rebuild -f -w winax');
  console.log('  - Then test from within Electron, not standalone Node');
  process.exit(1);
}

console.log('');

// Step 2: Check if GetObject exists (old API)
if (typeof winax.GetObject === 'function') {
  console.log('[INFO] winax.GetObject exists (older winax version)');
} else {
  console.log('[INFO] winax.GetObject does NOT exist');
  console.log('       Use: new winax.Object("Excel.Application", { activate: true })');
}

console.log('');

// Step 3: Try to connect to Excel using the CORRECT API
console.log('Attempting to connect to Excel...');
console.log('');

try {
  // This is the CORRECT way for winax 3.6.x
  // activate: true = connect to existing instance (like GetObject)
  const excel = new winax.Object('Excel.Application', { activate: true });

  console.log('[OK] Connected to Excel!');
  console.log('');
  console.log('Excel Properties:');
  console.log('  - Visible:', excel.Visible);
  console.log('  - Ready:', excel.Ready);
  console.log('  - Version:', excel.Version);
  console.log('  - WindowState:', excel.WindowState);

  // Try to get window position
  console.log('');
  console.log('Window Position (in points):');
  console.log('  - Left:', excel.Left);
  console.log('  - Top:', excel.Top);
  console.log('  - Width:', excel.Width);
  console.log('  - Height:', excel.Height);

  // Try to get active workbook
  const workbook = excel.ActiveWorkbook;
  if (workbook) {
    console.log('');
    console.log('Active Workbook:');
    console.log('  - Name:', workbook.Name);
    console.log('  - Path:', workbook.FullName);
  } else {
    console.log('');
    console.log('[WARN] No active workbook (open a file in Excel)');
  }

  // Clean up
  winax.release(excel);
  console.log('');
  console.log('[OK] Test completed successfully!');

} catch (err) {
  console.log('[FAIL] Could not connect to Excel');
  console.log('');
  console.log('Error Details:');
  console.log('  - Message:', err.message);
  console.log('  - Name:', err.name);
  if (err.stack) {
    console.log('  - Stack:', err.stack.split('\n').slice(0, 3).join('\n'));
  }
  console.log('');
  console.log('Troubleshooting:');
  console.log('  1. Is Excel running? (not just splash screen)');
  console.log('  2. Is Excel visible? (not minimized to tray)');
  console.log('  3. Are you running as Admin while Excel is not? (or vice versa)');
  console.log('  4. Kill all EXCEL.EXE processes and restart Excel fresh');
}

console.log('');
console.log('='.repeat(60));
