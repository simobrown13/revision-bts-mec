const fs = require('fs');
const os = require('os');
const pathMod = require('path');
const tmpDir = os.tmpdir();
const path = pathMod.join(tmpDir, 'pptx_extract', 'extracted', 'ppt', 'slides') + pathMod.sep;
const slides = [12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,29,30];

for (const i of slides) {
  const filePath = path + 'slide' + i + '.xml';
  if (!fs.existsSync(filePath)) {
    console.log('========== SLIDE ' + i + ' ========== FILE NOT FOUND');
    continue;
  }
  const xml = fs.readFileSync(filePath, 'utf-8');
  console.log('========== SLIDE ' + i + ' ==========');
  const spBlocks = xml.split('<p:sp>');
  for (let b = 1; b < spBlocks.length; b++) {
    const block = spBlocks[b].split('</p:sp>')[0];
    const blockTexts = [];
    const re2 = /<a:t>([^<]*)<\/a:t>/g;
    let m2;
    while ((m2 = re2.exec(block)) !== null) {
      blockTexts.push(m2[1]);
    }
    if (blockTexts.length > 0) {
      console.log(blockTexts.join(''));
      console.log('---');
    }
  }
  console.log('');
}
