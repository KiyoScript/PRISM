function pp_skylinePack(pieces, rollWidth, rollAvail, limitAvail) {
  var remaining = pieces.slice();
  var placed = [];
  var skyline = [{ x: 0, w: rollWidth, y: 0 }];
  var totalLen = 0;
  var totalPieceArea = 0;
  var unsupportedPieces = 0;

  // For visual grouping later, optionally we can keep track...
  
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
        
        // Find all possible X positions
        for (var s = 0; s < skyline.length; s++) {
          var startX = skyline[s].x;
          if (startX + o.w > rollWidth + 0.0001) continue;

          // Find max Y in [startX, startX + o.w]
          var maxY = 0;
          for (var js = 0; js < skyline.length; js++) {
            var sl = skyline[js];
            if (sl.x < startX + o.w - 0.001 && sl.x + sl.w > startX + 0.001) {
              if (sl.y > maxY) maxY = sl.y;
            }
          }

          // Evaluate this placement
          var score = maxY; 
          var tieBreaker = startX * 0.0001 - (o.w * o.h) * 0.000001; // Prefer lower Y, then left, then bigger items
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
      // Piece completely unplaceable (too wide in both dimensions)
      var forced = remaining.shift();
      forced.w = Math.min(forced.origW, forced.origH);
      forced.h = Math.max(forced.origW, forced.origH);
      forced.rotated = false;
      forced.oversized = true;
      forced.dx = 0;
      forced.dy = totalLen;
      
      if (limitAvail && totalLen > rollAvail + 0.0001) {
          remaining.unshift(forced);
          break; // Stop completely for actual packing
      }

      placed.push(forced);
      totalLen += forced.h;
      totalPieceArea += forced.w * forced.h;
      unsupportedPieces++;
      skyline = [{ x: 0, w: rollWidth, y: totalLen }];
      continue;
    }

    // Now check if this new placement exceeds avail
    var picked = remaining[bestIdx];
    if (limitAvail && bestY + bestOr.h > rollAvail + 0.0001) {
      // With Skyline, another smaller piece MIGHT still fit in the current valleys.
      // But typically if the best piece exceeds available length, we try other pieces?
      // Wait, the "best" piece (lowest score) exceeded bounds.
      break; 
    }

    // It fits! Remove from remaining
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

    // Update skyline
    var startX = bestX;
    var endX = startX + picked.w;
    var newSkyline = [];

    for (var s = 0; s < skyline.length; s++) {
      var sl = skyline[s];
      var slEnd = sl.x + sl.w;
      
      if (slEnd <= startX + 0.001 || sl.x >= endX - 0.001) {
        newSkyline.push(sl);
      } else {
        if (sl.x < startX) {
          newSkyline.push({ x: sl.x, y: sl.y, w: startX - sl.x });
        }
        if (slEnd > endX) {
          newSkyline.push({ x: endX, y: sl.y, w: slEnd - endX });
        }
      }
    }
    newSkyline.push({ x: startX, y: newY, w: picked.w });

    // Merge adjacent segments
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

  // Generate chunks/rows for rendering nicely
  // Instead of one huge row, we group things roughly by 'dy' or just return one row.
  // One row is the most honest representation of continuous packing.
  var totalAreaContext = rollWidth * totalLen;
  var wasteArea = Math.max(0, totalAreaContext - totalPieceArea);
  var wastePct = totalAreaContext > 0 ? Math.round(wasteArea / totalAreaContext * 100) : 0;
  
  var rows = [{
    pieces: placed,
    usedW: rollWidth,
    rowH: totalLen,
    wasteW: 0 // We don't have visual right-waste anymore, it's integrated
  }];

  return {
    rows: rows,
    totalLen: totalLen,
    remainingPieces: remaining,
    totalPieces: placed.length,
    wastePct: wastePct,
    unsupportedPieces: unsupportedPieces,
    placed: placed
  };
}
