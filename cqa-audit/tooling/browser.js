const { chromium } = require('/opt/node22/lib/node_modules/playwright');
const PROXY = process.env.HTTPS_PROXY || process.env.https_proxy;
async function launch() {
  return chromium.launch({
    executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome',
    proxy: PROXY ? { server: PROXY } : undefined,
    args: ['--no-sandbox','--disable-gpu','--ssl-version-max=tls1.2']
  });
}
module.exports = { launch };
