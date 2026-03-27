function buildVisualRows(placed, rollWidth, totalLen) {
  var ys = new Set([0, totalLen]);
  placed.forEach(function(p) {
    ys.add(p.dy);
    ys.add(p.dy + p.h);
  });
  var yArr = Array.from(ys).sort(function(a, b) { return a - b; });
  
  var rows = [];
  var currentPiecesForDisplay = [];
  
  for (var i = 0; i < yArr.length - 1; i++) {
    var y1 = yArr[i];
    var y2 = yArr[i+1];
    var h = y2 - y1;
    if (h <= 0.001) continue;
    
    var usedW = 0;
    placed.forEach(function(p) {
      // If piece overlaps this vertical band
      if (p.dy < y2 - 0.001 && p.dy + p.h > y1 + 0.001) {
        usedW += p.w;
      }
    });
    
    // We don't want to duplicate piece rendering in every band, 
    // so we assign pieces to the row where they START.
    var rowPieces = [];
    placed.forEach(function(p) {
      if (Math.abs(p.dy - y1) < 0.001) {
        // Clone for display coordinates relative to this row?
        // Wait, the UI uses p.dy as absolute or relative?
        // UI code: py = y + (p.dy !== undefined ? p.dy : 0);
        // BUT my pp_skylinePack sets p.dy as absolute! And then puts them in a row.
        // Wait, if p.dy is absolute, and UI does y += rowDispH, then py = y + p.dy?
        // Let's check PRISMPlotting.html pp_drawCanvas!
        rowPieces.push(p);
      }
    });
    
    rows.push({
      pieces: rowPieces,
      usedW: usedW,
      rowH: h,
      wasteW: Math.max(0, rollWidth - usedW)
    });
  }
  return rows;
}
