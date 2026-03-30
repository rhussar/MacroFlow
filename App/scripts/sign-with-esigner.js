const { execSync } = require('child_process');
const path = require('path');

const CODESIGNTOOL = process.env.CODESIGNTOOL_PATH || 'C:\\CodeSignTool\\CodeSignTool.bat';

const SIGNABLE_EXTENSIONS = new Set(['.exe', '.dll', '.msi', '.node']);

module.exports = async function sign(fileToSign) {
  const ext = path.extname(fileToSign).toLowerCase();
  if (!SIGNABLE_EXTENSIONS.has(ext)) return;

  const command = [
    `"${CODESIGNTOOL}"`,
    'sign',
    `-username="${process.env.ES_USERNAME}"`,
    `-password="${process.env.ES_PASSWORD}"`,
    `-credential_id="${process.env.ES_CREDENTIAL_ID}"`,
    `-totp_secret="${process.env.ES_TOTP_SECRET}"`,
    `-input_file_path="${fileToSign}"`,
    `-output_dir_path="${path.dirname(fileToSign)}"`,
  ].join(' ');

  console.log(`Signing: ${path.basename(fileToSign)}`);
  execSync(command, { stdio: 'inherit', timeout: 120_000 });
};
