const fs = require('fs');
let code = `function pp_orientationsForSpace(p, rollWidth, spaceLeft) {
  var out = [];
  if (p.origW <= rollWidth && p.origW <= spaceLeft) out.push({ w: p.origW, h: p.origH, rotated: false });
  if (true && p.origH <= rollWidth && p.origH <= spaceLeft && p.origH !== p.origW) out.push({ w: p.origH, h: p.origW, rotated: true });
  return out;
}`;
code += '\n' + fs.readFileSync('skyline_test.js', 'utf8');
code += `
var pieces = [];
for (var i=0; i<12; i++) pieces.push({ origW: 2.75, origH: 6.5, rotated: false });
var res = pp_skylinePack(pieces, 10.5, 100, false);
console.log('Total Len:', res.totalLen);
for (var i=0; i<res.placed.length; i++) {
  var p = res.placed[i];
  console.log('Placed ' + p.w + 'x' + p.h + ' at X:' + p.dx + ' Y:' + p.dy);
}
`;
fs.writeFileSync('run_skyline_2.js', code);