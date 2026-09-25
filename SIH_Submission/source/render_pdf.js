const { chromium } = require('playwright-core');
(async () => {
  const [src, out] = process.argv.slice(2);
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const p = await b.newPage();
  await p.goto('file://' + src, { waitUntil: 'networkidle' });
  await p.evaluate(() => document.fonts.ready);
  // overflow check
  const ov = await p.evaluate(() => [...document.querySelectorAll('.card,.bmc .b,.page')].filter(e => e.scrollHeight > e.clientHeight + 2).map(e => (e.className + ' | ' + (e.querySelector('h3,h2')||{}).textContent + ' ' + e.scrollHeight + '>' + e.clientHeight)));
  console.log('OVERFLOW:', JSON.stringify(ov, null, 1));
  await p.pdf({ path: out, width: '297mm', height: '210mm', printBackground: true, preferCSSPageSize: true });
  await b.close();
})();
