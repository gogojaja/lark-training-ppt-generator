const { execFile } = require('child_process');
const path = require('path');

function isLarkCliAvailable() {
  return new Promise((resolve) => {
    execFile('lark-cli', ['--version'], { windowsHide: true }, (error) => {
      resolve(!error);
    });
  });
}

async function exportToLark(outputFile, options = {}) {
  const filePath = path.resolve(outputFile);
  const localOnly = !!options.localOnly;

  if (localOnly) {
    return {
      ok: false,
      skipped: true,
      reason: 'local-only mode enabled; cloud export disabled',
    };
  }

  const available = await isLarkCliAvailable();
  if (!available) {
    return {
      ok: false,
      skipped: true,
      reason: 'lark-cli not installed; cloud export is optional and not required for local generation',
    };
  }

  return new Promise((resolve) => {
    execFile(
      'lark-cli',
      ['docs', '+upload', '--file', filePath],
      { windowsHide: true },
      (error, stdout, stderr) => {
        if (error) {
          resolve({
            ok: false,
            skipped: true,
            reason: stderr || error.message || 'lark-cli upload failed',
          });
          return;
        }

        resolve({
          ok: true,
          skipped: false,
          output: stdout || 'Lark upload command executed successfully',
        });
      }
    );
  });
}

module.exports = {
  exportToLark,
  isLarkCliAvailable,
};
