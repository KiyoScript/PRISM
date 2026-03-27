function pp_computeWasteRects(placed, rollWidth, totalLen) {
    if (totalLen <= 0) return [];
    var step = 4; // 1/4 ft resolution is usually safe for imperial dimensions here?
    // Let's use 10 for safe decimals? Pieces are 2.75, 6.5. 2.75 * 4 = 11. exact.
    // 100 is even safer.
    var res = 100;
    var W = Math.round(rollWidth * res);
    var H = Math.round(totalLen * res);
    
    // Using a 1D array instead of grid since W is small
    // Actually, sweeping rectangles is exact and fast.
    // How to sweep rectangles:
    var rects = [];
    var yBreaks = new Set([0, totalLen]);
    placed.forEach(function(p) {
        yBreaks.add(p.dy);
        yBreaks.add(p.dy + p.h);
    });
    var yArr = Array.from(yBreaks).sort(function(a, b) { return a - b; });

    for (var i = 0; i < yArr.length - 1; i++) {
        var y1 = yArr[i];
        var y2 = yArr[i+1];
        if (y2 - y1 < 0.01) continue;

        var xBreaks = new Set([0, rollWidth]);
        placed.forEach(function(p) {
            if (p.dy < y2 - 0.001 && p.dy + p.h > y1 + 0.001) {
                xBreaks.add(p.dx);
                xBreaks.add(p.dx + p.w);
            }
        });
        var xArr = Array.from(xBreaks).sort(function(a, b) { return a - b; });

        for (var j = 0; j < xArr.length - 1; j++) {
            var x1 = xArr[j];
            var x2 = xArr[j+1];
            if (x2 - x1 < 0.01) continue;

            // Is this cell occupied?
            var cx = (x1 + x2) / 2;
            var cy = (y1 + y2) / 2;
            var occupied = false;
            for (var p = 0; p < placed.length; p++) {
                var pc = placed[p];
                if (pc.dx < cx && pc.dx + pc.w > cx && pc.dy < cy && pc.dy + pc.h > cy) {
                    occupied = true;
                    break;
                }
            }
            if (!occupied) {
                rects.push({ dx: x1, dy: y1, w: x2 - x1, h: y2 - y1 });
            }
        }
    }
    
    // Merge vertically
    var merged = [];
    rects.forEach(function(r) {
        var merged_it = false;
        for (var k = 0; k < merged.length; k++) {
            var m = merged[k];
            if (Math.abs(m.dx - r.dx) < 0.001 && Math.abs(m.w - r.w) < 0.001 && Math.abs(m.dy + m.h - r.dy) < 0.001) {
                m.h += r.h;
                merged_it = true;
                break;
            }
        }
        if (!merged_it) merged.push(r);
    });

    return merged;
}
module.exports = pp_computeWasteRects;
