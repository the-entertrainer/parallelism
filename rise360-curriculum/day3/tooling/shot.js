const { chromium } = require('/opt/node22/lib/node_modules/playwright');
const path = require('path');
(async () => {
  const b = await chromium.launch({
    executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome',
    args: ['--no-sandbox','--disable-gpu','--ssl-version-max=tls1.2'],
  });
  const p = await b.newPage({ viewport: { width: 640, height: 700 }, deviceScaleFactor: 2 });
  await p.goto('file://' + path.join(__dirname, '..', 'bonus', 'swipe-deck.html'));
  await p.waitForTimeout(700);

  // resting stack
  await p.screenshot({ path: '../assets/swipe-screenshots/swipe_rest.png' });

  // mid-drag: press on the top card and move right past half the threshold
  const card = await p.$('.card');
  const box = await card.boundingBox();
  const cx = box.x + box.width / 2, cy = box.y + box.height / 2;
  await p.mouse.move(cx, cy);
  await p.mouse.down();
  await p.mouse.move(cx + 58, cy - 10, { steps: 12 });
  await p.waitForTimeout(220);
  await p.screenshot({ path: '../assets/swipe-screenshots/swipe_drag.png' });
  await p.mouse.up();
  await p.waitForTimeout(600);

  // finished state
  for (let i = 0; i < 6; i++) { await p.keyboard.press('ArrowRight'); await p.waitForTimeout(420); }
  await p.waitForTimeout(400);
  await p.screenshot({ path: '../assets/swipe-screenshots/swipe_done.png' });

  console.log('shots done');
  await b.close();
})();
