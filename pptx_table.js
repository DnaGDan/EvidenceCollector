/**
 * addBottomRightTable + helpers for pptxgenjs (v3.12+)
 *
 * - Draws a crisp editable grid in the bottom-right of a slide using shapes + text.
 * - No PPT "table" widget; manual lines/rects and text only.
 * - All geometry in inches. Text measurement in px (1in = 96px). 1pt = 1/72in -> px = pt * 96/72.
 *
 * Exports:
 *  - addBottomRightTable(slide, opts)
 *  - measureTextPx(text, fontPx, bold)
 *  - wrapWithEllipsis(text, fontPx, maxWidthPx, maxLines)
 *  - downloadPptx(pptx, filename)
 *
 * Usage snippet at bottom.
 */

/* -------------------- Small constants & cached canvas -------------------- */
const _DPI = 96;                 // px per inch
const _PT_TO_PX = _DPI / 72;     // 1pt -> px
const _BASE_GRID_IN = 0.125;     // baseline grid (inches)
let _measureCtx = null;
function _getMeasureCtx(){
  if(_measureCtx) return _measureCtx;
  const c = document.createElement('canvas');
  _measureCtx = c.getContext('2d');
  return _measureCtx;
}

/* -------------------- Text measure & wrap helpers -------------------- */
/**
 * measureTextPx(text, fontPx, bold = false)
 * Returns width in pixels for Arial.
 */
export function measureTextPx(text, fontPx, bold = false){
  try{
    const ctx = _getMeasureCtx();
    ctx.font = `${bold ? 'bold ' : ''}${Math.round(fontPx)}px Arial`;
    return ctx.measureText(text || '').width;
  }catch(e){
    console.warn('measureTextPx failed', e);
    return (text||'').length * (fontPx * 0.5);
  }
}

/**
 * wrapWithEllipsis(text, fontPx, maxWidthPx, maxLines)
 * - Returns array of lines (<= maxLines).
 * - Manual word-wrap; breaks long tokens; final line may include '…'.
 */
export function wrapWithEllipsis(text, fontPx, maxWidthPx, maxLines=2){
  if(!text) return [];
  const words = String(text).split(/\s+/);
  const lines = [];
  let cur = '';

  const ell = '…';
  const spaceWidth = measureTextPx(' ', fontPx, false);

  for(let i=0;i<words.length;i++){
    const w = words[i];
    // measure token
    const tokW = measureTextPx(w, fontPx, false);

    if(!cur){
      // if single token wider than maxWidthPx -> hard break that token
      if(tokW > maxWidthPx){
        // break token into chars
        let partial = '';
        for(const ch of w){
          const withCh = partial + ch;
          if(measureTextPx(withCh, fontPx, false) <= maxWidthPx){
            partial = withCh;
          }else{
            lines.push(partial);
            partial = ch;
            if(lines.length >= maxLines) break;
          }
        }
        if(partial && lines.length < maxLines) cur = partial;
      }else{
        cur = w;
      }
    }else{
      // try append with space
      const newLine = cur + ' ' + w;
      if(measureTextPx(newLine, fontPx, false) <= maxWidthPx){
        cur = newLine;
      }else{
        lines.push(cur);
        cur = '';
        // reprocess this word next iteration (i--), but handle oversized token
        i--;
      }
    }
    if(lines.length >= maxLines) break;
  }

  if(lines.length < maxLines && cur) lines.push(cur);

  // If still more content (words left), truncate last line with ellipsis
  if(lines.length > 0 && (words.length > 0) ){
    // Determine if original text fit completely
    const recomposed = lines.join(' ') + (cur ? ' ' + cur : '');
    if(recomposed.trim().length < text.trim().length && lines.length > 0){
      // Need to ellipsize last line to fit
      let last = lines[lines.length - 1];
      const maxW = maxWidthPx;
      if(measureTextPx(last + ell, fontPx, false) > maxW){
        // trim chars until fits
        while(last.length && measureTextPx(last + ell, fontPx, false) > maxW){
          last = last.slice(0, -1);
        }
      }
      lines[lines.length - 1] = last + ell;
    } else {
      // Edge-case: very long single token that spilled into more lines than allowed
      if(lines.length === maxLines){
        const lastIdx = lines.length - 1;
        let last = lines[lastIdx];
        if(measureTextPx(last, fontPx, false) > maxWidthPx){
          while(last.length && measureTextPx(last + ell, fontPx, false) > maxWidthPx){
            last = last.slice(0, -1);
          }
          lines[lastIdx] = last + ell;
        }
      }
    }
  }

  // Ensure not exceeding maxLines
  return lines.slice(0, maxLines);
}

/* -------------------- Core: addBottomRightTable -------------------- */
/**
 * addBottomRightTable(slide, opts)
 *
 * opts:
 *  - slideWidthIn, slideHeightIn, marginIn,
 *  - areaWidthIn, areaHeightIn, areaRightIn, areaBottomIn,
 *  - cols, rows,
 *  - strokePt = 1, labelPt=10, valuePt=12, cellPaddingIn=0.08,
 *  - palette line/label/text/fill,
 *  - data: TableCell[][] rows x cols. TableCell: { label?, value, maxLines?, fontPt? }
 */
export async function addBottomRightTable(slide, opts){
  if(!slide || typeof slide.addText !== 'function') throw new Error('Invalid slide object (PptxGenJS slide required).');

  // defaults
  const o = Object.assign({
    strokePt: 1,
    labelPt: 10,
    valuePt: 12,
    cellPaddingIn: 0.08,
    palette: { line: '7A93A6', label: '6B7280', text: '000000', fill: 'FFFFFF' }
  }, opts);

  const {
    slideWidthIn, slideHeightIn, marginIn,
    areaWidthIn, areaHeightIn, areaRightIn, areaBottomIn,
    cols, rows, strokePt, labelPt, valuePt, cellPaddingIn, palette
  } = o;

  // compute area top-left so area is bottom-right aligned to given distances
  // x = slideWidth - areaRight - areaWidth
  const areaX = (slideWidthIn - areaRightIn - areaWidthIn);
  const areaY = (slideHeightIn - areaBottomIn - areaHeightIn);

  // cell sizes (we don't subtract stroke thickness from geometry — lines drawn on edges)
  const cellW = areaWidthIn / cols;
  const cellH = areaHeightIn / rows;

  // precompute px sizes for text measurement
  const labelFontPx = labelPt * _PT_TO_PX;
  const valueFontPx = valuePt * _PT_TO_PX;

  // Draw outer rect for area (stroke outlines the whole table)
  try{
    slide.addShape(pptx.ShapeType.rect, {
      x: areaX, y: areaY, w: areaWidthIn, h: areaHeightIn,
      fill: { color: palette.fill },
      line: { color: palette.line, width: strokePt }
    });
  }catch(e){
    // If addShape rect fails, try without fill
    slide.addShape(pptx.ShapeType.rect, { x: areaX, y: areaY, w: areaWidthIn, h: areaHeightIn, line:{ color: palette.line, width: strokePt } });
  }

  // Internal vertical lines (cols-1)
  for(let c=1;c<cols;c++){
    const x = areaX + c * cellW;
    slide.addShape(pptx.ShapeType.line, {
      x: x, y: areaY, w: 0, h: areaHeightIn,
      line: { color: palette.line, width: strokePt }
    });
  }
  // Internal horizontal lines (rows-1)
  for(let r=1;r<rows;r++){
    const y = areaY + r * cellH;
    slide.addShape(pptx.ShapeType.line, {
      x: areaX, y: y, w: areaWidthIn, h: 0,
      line: { color: palette.line, width: strokePt }
    });
  }

  // For each cell: draw inner rect (no double outer strokes) and add label+value
  for(let r=0;r<rows;r++){
    for(let c=0;c<cols;c++){
      const cellX = areaX + c * cellW;
      const cellY = areaY + r * cellH;

      // label height reserve (inches). Convert labelPt -> inches approx: labelPt/72
      const labelHeightIn = (labelPt / 72) * 1.1; // 10% breathing room
      // compute box (rect) area inside cell below the label
      const boxX = cellX + cellPaddingIn;
      const boxW = cellW - cellPaddingIn * 2;
      const boxY = cellY + labelHeightIn + (cellPaddingIn / 2);
      const boxH = cellH - labelHeightIn - cellPaddingIn;

      // draw the cell inner rectangle (fill white, thin border)
      slide.addShape(pptx.ShapeType.rect, {
        x: boxX, y: boxY, w: boxW, h: boxH,
        fill: { color: palette.fill },
        line: { color: palette.line, width: strokePt * 0.7 }
      });

      // fetch cell data if provided
      const cell = (o.data && o.data[r] && o.data[r][c]) ? o.data[r][c] : { value: '' };
      const labelText = cell.label ? String(cell.label) : '';
      const valueText = cell.value ? String(cell.value) : '';
      const maxLines = typeof cell.maxLines === 'number' ? cell.maxLines : 2;
      const cellValuePt = typeof cell.fontPt === 'number' ? cell.fontPt : valuePt;
      const cellValuePx = cellValuePt * _PT_TO_PX;

      // Draw label (above the box, small)
      if(labelText){
        const labelX = cellX + cellPaddingIn;
        const labelY = cellY + (cellPaddingIn / 4);
        // We do not allow label to wrap multiple lines; truncate with ellipsis if needed
        const maxLabelWidthPx = Math.round((boxW - (cellPaddingIn*0)) * _DPI);
        const labelLines = wrapWithEllipsis(labelText, labelFontPx, maxLabelWidthPx, 1);
        slide.addText(labelLines.join('\n'), {
          x: labelX, y: labelY, w: boxW, h: labelHeightIn,
          fontSize: labelPt, color: palette.label, fontFace: 'Arial', align: 'left', valign: 'top'
        });
      }

      // Prepare value lines: measure available text width inside box minus small padding
      const textPaddingPx = Math.round((cellPaddingIn * _DPI) * 0.5); // a bit of inner padding in px
      const maxWidthPx = Math.max(8, Math.round(boxW * _DPI) - (textPaddingPx * 2));
      const wrapped = wrapWithEllipsis(valueText, cellValuePx, maxWidthPx, maxLines);

      if(wrapped.length){
        // Create final text with explicit line breaks, manual font and no autosize
        slide.addText(wrapped.join('\n'), {
          x: boxX + (cellPaddingIn / 2), y: boxY + (cellPaddingIn / 2),
          w: boxW - (cellPaddingIn), h: boxH - (cellPaddingIn),
          fontSize: cellValuePt, color: palette.text, fontFace: 'Arial', align: 'left', valign: 'top'
        });
      }
    }
  }
}

/* -------------------- download helper for pptxgenjs -------------------- */
/**
 * downloadPptx(pptx, filename)
 * - Writes arraybuffer, wraps to PPTX mime, and triggers download anchor click.
 */
export async function downloadPptx(pptx, filename = 'presentation.pptx'){
  if(typeof PptxGenJS === 'undefined' || !pptx || typeof pptx.write !== 'function'){
    alert('PptxGenJS not available or invalid presentation object.');
    return;
  }
  try{
    const ab = await pptx.write('arraybuffer');
    const pptxMime = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
    const blob = new Blob([ab], { type: pptxMime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 3000);
  }catch(err){
    console.error('downloadPptx failed', err);
    alert('Failed to generate PPTX: ' + (err && err.message ? err.message : String(err)));
  }
}

/* -------------------- Usage snippet --------------------
   Example: create pptx, slide, add table, download.
   (Drop this into your code where you have access to PptxGenJS.)
------------------------------------------------------------------ */
(function usageExample(){
  // guard: only run when explicitly imported as module? Wrap so it doesn't auto-run in production.
  if(window.__pptx_table_usage_run__) return;
  window.__pptx_table_usage_run__ = false; // set to true to allow demo run

  // If you want to test manually, set the flag to true in console:
  // window.__pptx_table_usage_run__ = true; then call window.__pptx_table_demo()
  window.__pptx_table_demo = async function(){
    if(typeof PptxGenJS === 'undefined'){ alert('PptxGenJS missing'); return; }

    const pptx = new PptxGenJS();
    // Make sure A4 layout exists (example: landscape A4)
    try{
      pptx.defineLayout({ name:'A4LAND', width:11.69, height:8.27 });
      pptx.layout = 'A4LAND';
    }catch(e){ /* ignore */ }

    const slide = pptx.addSlide();

    // Example table data (2 rows x 3 cols), with a pathological token
    const data = [
      [
        { label: 'Borehole ID', value: 'BH-00123-VERY-LONG-TOKEN-THAT-NEEDS-HARD-BREAK-ABCDEFGHIJKLMNOPQRSTUVWXYZ', maxLines: 3, fontPt: 10 },
        { label: 'Chainage', value: 'Chainage 1234.56 m, long description continues to test wrapping and ellipsis behaviour', maxLines: 2 },
        { label: 'Material', value: 'Silty sand with gravel and occasional organic lenses. Very long test string to wrap.' }
      ],
      [
        { label: 'Compaction', value: '95% (tested on site)' },
        { label: 'Notes', value: 'Observed wet patch near NW corner. Contractor to review.' },
        { label: 'Inspector', value: 'J. Example, MSc, CGeol' }
      ]
    ];

    // Provide options matching your slide size & margins
    await addBottomRightTable(slide, {
      slideWidthIn: 11.69,
      slideHeightIn: 8.27,
      marginIn: 0.35,
      areaWidthIn: 2.1,      // width of bottom-right framework (inches)
      areaHeightIn: 3.2,     // height (inches)
      areaRightIn: 0.35,     // distance from right edge (inches)
      areaBottomIn: 0.35,    // distance from bottom edge (inches)
      cols: 3,
      rows: 2,
      strokePt: 1,
      labelPt: 9,
      valuePt: 10,
      cellPaddingIn: 0.08,
      palette: { line: '7A93A6', label: '6B7280', text: '000000', fill: 'FFFFFF' },
      data
    });

    // then download
    await downloadPptx(pptx, 'evidence_table_demo.pptx');
  };
})();

/* Expose to global for non-module usage */
window.addBottomRightTable = addBottomRightTable;
window.measureTextPx = measureTextPx;
window.wrapWithEllipsis = wrapWithEllipsis;
window.downloadPptx = downloadPptx;