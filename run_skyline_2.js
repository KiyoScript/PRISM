function pp_orientationsForSpace(p, rollWidth, spaceLeft) {
  var out = [];
  if (p.origW <= rollWidth && p.origW <= spaceLeft) out.push({ w: p.origW, h: p.origH, rotated: false });
  if (true && p.origH <= rollWidth && p.origH <= spaceLeft && p.origH !== p.origW) out.push({ w: p.origH, h: p.origW, rotated: true });
  return out;
}
function pp_skylinePack(pieces, rollWidth, rollAvail, limitAvail) {
  var remaining = pieces.slice();
  var placed = [];
  var skyline = [{ x: 0, w: rollWidth, y: 0 }];
  var totalLen = 0;
  var totalPieceArea = 0;
  var unsupportedPieces = 0;

  while (remaining.length > 0) {
    var bestIdx = -1;
    var bestOr = null;
    var bestX = 0;
    var bestY = 9999999;
    var bestScore = 9999999;

    for (var i = 0; i < remaining.length; i++) {
        var p = remaining[i];
        var ors = pp_orientationsForSpace(p, rollWidth, rollWidth);
        for (var oi = 0; oi < ors.length; oi++) {
            var o = ors[oi];
            for (var s = 0; s < skyline.length; s++) {
                var startX = skyline[s].x;
                if (startX + o.w > rollWidth + 0.0001) continue;
                var maxY = 0;
                for (var js = 0; js < skyline.length; js++) {
                    var sl = skyline[js];
                    if (sl.x < startX + o.w - 0.001 && sl.x + sl.w > startX + 0.001) {
                        if (sl.y > maxY) maxY = sl.y;
                    }
                }
                var increaseInLen = Math.max(0, (maxY + o.h) - totalLen);
                var score = increaseInLen * 1000 + maxY;
                var tieBreaker = startX * 0.0001 - (o.w * o.h) * 0.000001; 
                var totalScore = score + tieBreaker;

                if (totalScore < bestScore) {
                    bestScore = totalScore;
                    bestIdx = i;
                    bestOr = o;
                    bestX = startX;
                    bestY = maxY;
                }
            }
        }
    }

    if (bestIdx === -1) {
      var forced = remaining.shift();
      forced.w = Math.min(forced.origW, forced.origH);
      forced.h = Math.max(forced.origW, forced.origH);
      forced.rotated = false;
      forced.oversized = true;
      forced.dx = 0;
      forced.dy = totalLen;
      if (limitAvail && totalLen > rollAvail + 0.0001) { remaining.unshift(forced); break; }
      placed.push(forced);
      totalLen += forced.h;
      totalPieceArea += forced.w * forced.h;
      unsupportedPieces++;
      skyline = [{ x: 0, w: rollWidth, y: totalLen }];
      continue;
    }

    var picked = remaining[bestIdx];
    if (limitAvail && bestY + bestOr.h > rollAvail + 0.0001) { break; }

    picked = remaining.splice(bestIdx, 1)[0];
    picked.w = bestOr.w;
    picked.h = bestOr.h;
    picked.rotated = bestOr.rotated;
    picked.oversized = false;
    picked.dx = bestX;
    picked.dy = bestY;
    placed.push(picked);
    totalPieceArea += picked.w * picked.h;

    var newY = bestY + picked.h;
    if (newY > totalLen) totalLen = newY;

    var startX = bestX;
    var endX = startX + picked.w;
    var newSkyline = [];
    for (var s = 0; s < skyline.length; s++) {
      var sl = skyline[s];
      var slEnd = sl.x + sl.w;
      if (slEnd <= startX + 0.001 || sl.x >= endX - 0.001) {
        newSkyline.push(sl);
      } else {
        if (sl.x < startX) { newSkyline.push({ x: sl.x, y: sl.y, w: startX - sl.x }); }
        if (slEnd > endX) { newSkyline.push({ x: endX, y: sl.y, w: slEnd - endX }); }
      }
    }
    newSkyline.push({ x: startX, y: newY, w: picked.w });
    newSkyline.sort(function(a, b) { return a.x - b.x; });
    var merged = [];
    for (var m = 0; m < newSkyline.length; m++) {
      var nsl = newSkyline[m];
      if (merged.length > 0) {
        var last = merged[merged.length - 1];
        if (Math.abs(last.y - nsl.y) < 0.001 && Math.abs(last.x + last.w - nsl.x) < 0.001) {
          last.w += nsl.w;
          continue;
        }
      }
      merged.push(nsl);
    }
    skyline = merged;
  }

  return {
    rows: [{ pieces: placed, usedW: rollWidth, rowH: totalLen, wasteW: 0 }],
    totalLen: totalLen,
    remainingPieces: remaining,
    placed: placed
  };
}

var pieces = [];
for (var i=0; i<12; i++) pieces.push({ origW: 2.75, origH: 6.5, rotated: false });
var res = pp_skylinePack(pieces, 10.5, 100, false);
console.log('Total Len:', res.totalLen);
for (var i=0; i<res.placed.length; i++) {
  var p = res.placed[i];
  console.log('Placed ' + p.w + 'x' + p.h + ' at X:' + p.dx + ' Y:' + p.dy);
}
