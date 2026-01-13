/**
 * addBottomRightTable + helpers for pptxgenjs (v3.12+)
 *
 * Supports colspan/rowspan and optional imageData per cell.
 *
 * Exports:
 *  - addBottomRightTable(slide, opts)
 *  - measureTextPx(text, fontPx, bold)
 *  - wrapWithEllipsis(text, fontPx, maxWidthPx, maxLines)
 *  - downloadPptx(pptx, filename)
 */

/* -------------------- Small constants & cached canvas -------------------- */
const _DPI = 96;                 // px per inch
const _PT_TO_PX = _DPI / 72;     // 1pt -> px
let _measureCtx = null;
function _getMeasureCtx(){
  if(_measureCtx) return _measureCtx;
  const c = document.createElement('canvas');
  _measureCtx = c.getContext('2d');
  return _measureCtx;
}

/* -------------------- Text measure & wrap helpers -------------------- */
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

export function wrapWithEllipsis(text, fontPx, maxWidthPx, maxLines=2){
  if(!text) return [];
  const words = String(text).split(/\s+/);
  const lines = [];
  let cur = '';

  const ell = '…';

  for(let i=0;i<words.length;i++){
    const w = words[i];
    const tokW = measureTextPx(w, fontPx, false);

    if(!cur){
      if(tokW > maxWidthPx){
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
      const newLine = cur + ' ' + w;
      if(measureTextPx(newLine, fontPx, false) <= maxWidthPx){
        cur = newLine;
      }else{
        lines.push(cur);
        cur = '';
        i--;
      }
    }
    if(lines.length >= maxLines) break;
  }

  if(lines.length < maxLines && cur) lines.push(cur);

  // Ellipsize last line if content remains
  const combinedLen = lines.join(' ').length;
  if(lines.length && combinedLen < String(text).trim().length){
    let last = lines[lines.length - 1];
    while(last.length && measureTextPx(last + ell, fontPx, false) > maxWidthPx){
      last = last.slice(0, -1);
    }
    lines[lines.length - 1] = last + ell;
  } else if(lines.length === maxLines){
    let last = lines[lines.length - 1];
    if(measureTextPx(last, fontPx, false) > maxWidthPx){
      while(last.length && measureTextPx(last + ell, fontPx, false) > maxWidthPx){
        last = last.slice(0, -1);
      }
      lines[lines.length - 1] = last + ell;
    }
  }

  return lines.slice(0, maxLines);
}

/* -------------------- Core: addBottomRightTable with colspan/rowspan -------------------- */
export async function addBottomRightTable(slide, opts){
  if(!slide || typeof slide.addText !== 'function') throw new Error('Invalid slide object (PptxGenJS slide required).');

  const defaults = {
    strokePt: 1,
    labelPt: 10,
    valuePt: 12,
    cellPaddingIn: 0.08,
    palette: { line: '7A93A6', label: '6B7280', text: '000000', fill: 'FFFFFF' }
  };
  const o = Object.assign({}, defaults, opts);
  const {
    slideWidthIn, slideHeightIn, marginIn,
    areaWidthIn, areaHeightIn, areaRightIn, areaBottomIn,
    cols, rows, strokePt, labelPt, valuePt, cellPaddingIn, palette
  } = o;

  const areaX = (slideWidthIn - areaRightIn - areaWidthIn);
  const areaY = (slideHeightIn - areaBottomIn - areaHeightIn);

  const cellW = areaWidthIn / cols;
  const cellH = areaHeightIn / rows;

  // occupancy map & anchor list for spans
  const occupied = Array.from({length: rows}, ()=> Array(cols).fill(false));
  const anchors = [];
  for(let r=0;r<rows;r++){
    for(let c=0;c<cols;c++){
      if(occupied[r][c]) continue;
      const cell = (o.data && o.data[r] && o.data[r][c]) ? o.data[r][c] : null;
      const colspan = (cell && Number(cell.colspan)>0) ? Math.min(cols - c, Number(cell.colspan)) : 1;
      const rowspan = (cell && Number(cell.rowspan)>0) ? Math.min(rows - r, Number(cell.rowspan)) : 1;
      for(let rr=r; rr<r+rowspan; rr++){
        for(let cc=c; cc<c+colspan; cc++){
          if(rr < rows && cc < cols) occupied[rr][cc] = true;
        }
      }
      anchors.push({ r, c, cell, colspan, rowspan });
    }
  }

  // outer rect
  try{
    slide.addShape(pptx.ShapeType.rect, {
      x: areaX, y: areaY, w: areaWidthIn, h: areaHeightIn,
      fill: { color: palette.fill },
      line: { color: palette.line, width: strokePt }
    });
  }catch(e){
    slide.addShape(pptx.ShapeType.rect, { x: areaX, y: areaY, w: areaWidthIn, h: areaHeightIn, line:{ color: palette.line, width: strokePt } });
  }

  // vertical lines (skip where colspan crosses)
  for(let colB = 1; colB < cols; colB++){
    let skip = false;
    for(const a of anchors){
      const start = a.c;
      const end = a.c + a.colspan;
      if(a.cell && (start < colB && end > colB)){ skip = true; break; }
    }
    if(!skip){
      const x = areaX + colB * cellW;
      slide.addShape(pptx.ShapeType.line, { x: x, y: areaY, w: 0, h: areaHeightIn, line: { color: palette.line, width: strokePt } });
    }
  }

  // horizontal lines (skip where rowspan crosses)
  for(let rowB = 1; rowB < rows; rowB++){
    let skip = false;
    for(const a of anchors){
      const start = a.r;
      const end = a.r + a.rowspan;
      if(a.cell && (start < rowB && end > rowB)){ skip = true; break; }
    }
    if(!skip){
      const y = areaY + rowB * cellH;
      slide.addShape(pptx.ShapeType.line, { x: areaX, y: y, w: areaWidthIn, h: 0, line: { color: palette.line, width: strokePt } });
    }
  }

  const labelFontPx = labelPt * _PT_TO_PX;
  const baseValueFontPx = valuePt * _PT_TO_PX;

  for(const a of anchors){
    const { r, c, cell, colspan, rowspan } = a;
    const x = areaX + c * cellW;
    const y = areaY + r * cellH;
    const w = colspan * cellW;
    const h = rowspan * cellH;

    const labelHeightIn = (labelPt / 72) * 1.1;
    const boxX = x + cellPaddingIn;
    const boxW = Math.max(0.01, w - cellPaddingIn*2);
    const boxY = y + labelHeightIn + (cellPaddingIn / 2);
    const boxH = Math.max(0.01, h - labelHeightIn - cellPaddingIn);

    slide.addShape(pptx.ShapeType.rect, {
      x: boxX, y: boxY, w: boxW, h: boxH,
      fill: { color: palette.fill },
      line: { color: palette.line, width: strokePt * 0.8 }
    });

    if(!cell) continue;

    const labelText = cell.label ? String(cell.label) : '';
    const valueText = cell.value ? String(cell.value) : '';
    const maxLines = typeof cell.maxLines === 'number' ? cell.maxLines : 2;
    const cellValuePt = typeof cell.fontPt === 'number' ? cell.fontPt : valuePt;
    const cellValuePx = cellValuePt * _PT_TO_PX;

    if(labelText){
      const maxLabelWidthPx = Math.round(boxW * _DPI);
      const lab = wrapWithEllipsis(labelText, labelFontPx, maxLabelWidthPx, 1);
      slide.addText(lab.join('\n'), {
        x: x + cellPaddingIn, y: y + (cellPaddingIn / 4), w: boxW, h: labelHeightIn,
        fontSize: labelPt, color: palette.label, fontFace: 'Arial', align: 'left', valign: 'top'
      });
    }

    if(cell.imageData){
      try{
        const imgPad = cellPaddingIn * 0.6;
        slide.addImage({
          data: cell.imageData,
          x: boxX + imgPad, y: boxY + imgPad,
          w: Math.max(0.01, boxW - imgPad*2),
          h: Math.max(0.01, boxH - imgPad*2)
        });
      }catch(e){
        console.warn('addBottomRightTable: image embed failed for cell', e);
      }
    }

    if(valueText){
      const textPaddingPx = Math.round((cellPaddingIn * _DPI) * 0.5);
      const maxWidthPx = Math.max(8, Math.round(boxW * _DPI) - (textPaddingPx * 2));
      const wrapped = wrapWithEllipsis(valueText, cellValuePx, maxWidthPx, maxLines);
      if(wrapped.length){
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

/* -------------------- Helper to generate opts.data for the Brownfield Solutions layout -------------------- */
/**
 * Pass imageData (base64) for the logo cell.
 */
export function getBrownfieldTableOpts({ client, projectTitle, title, docRef, page, pageCount, compiledBy, checkedBy, logoImageData }) {
  return {
    slideWidthIn: 10, slideHeightIn: 7.5,
    areaWidthIn: 4, areaHeightIn: 5.5,
    areaRightIn: 0.5, areaBottomIn: 0.5,
    cols: 5, rows: 10,
    data: [
      // Row 0: Header row
      [
        { label: 'REV', value: '', fontPt: 10 },
        { label: 'DATE', value: '', fontPt: 10 },
        { label: 'DESCRIPTION', value: '', fontPt: 10, colspan: 2 },
        { label: 'BY', value: '', fontPt: 10 },
        { label: 'CKD', value: '', fontPt: 10 }
      ],
      // Row 1: Empty row
      [ {}, {}, {}, {}, {} ],
      // Row 2: Empty row
      [ {}, {}, {}, {}, {} ],
      // Row 3: Logo row (spans all columns)
      [
        { imageData: logoImageData, colspan: 5, rowspan: 2 }
      ],
      // Row 4: (covered by logo rowspan)
      [ {}, {}, {}, {}, {} ],
      // Row 5: CLIENT row (spans all columns)
      [
        { label: 'CLIENT', value: client, colspan: 5 }
      ],
      // Row 6: PROJECT TITLE row (spans all columns)
      [
        { label: 'PROJECT TITLE', value: projectTitle, colspan: 5 }
      ],
      // Row 7: TITLE row (spans all columns)
      [
        { label: 'TITLE', value: title, colspan: 5 }
      ],
      // Row 8: Associated Doc Ref & Page
      [
        { label: 'Associated Document Ref No.', value: docRef, colspan: 2 },
        { label: 'PAGE', value: `${page} of ${pageCount}`, colspan: 3 }
      ],
      // Row 9: Compiled By & Checked By
      [
        { label: 'COMPLIED BY', value: compiledBy, colspan: 2 },
        { label: 'CHECKED BY', value: checkedBy, colspan: 3 }
      ]
    ]
  };
}

/**
 * Example usage:
 * const opts = getBrownfieldTableOpts({
 *   client: 'XXXX',
 *   projectTitle: 'XXXX',
 *   title: 'XXXX',
 *   docRef: 'CXXXX',
 *   page: 1,
 *   pageCount: 5,
 *   compiledBy: 'XX',
 *   checkedBy: 'XXX',
 *   logoImageData: 'data:image/png;base64,...'
 * });
 * await addBottomRightTable(slide, opts);
 */

/* Example: Attach to a button with id="exportBtn" */
document.getElementById('exportBtn').addEventListener('click', async function() {
  const opts = getBrownfieldTableOpts({
    client: 'XXXX',
    projectTitle: 'XXXX',
    title: 'XXXX',
    docRef: 'CXXXX',
    page: 1,
    pageCount: 5,
    compiledBy: 'XX',
    checkedBy: 'XXX',
    logoImageData: 'data:image/png;base64,...'
  });
  const pptx = new PptxGenJS();
  const slide = pptx.addSlide();
  await addBottomRightTable(slide, opts);
  await downloadPptx(pptx, 'BrownfieldTable.pptx');
});

/* Expose to window for legacy callers */
window.addBottomRightTable = addBottomRightTable;
window.measureTextPx = measureTextPx;
window.wrapWithEllipsis = wrapWithEllipsis;
window.downloadPptx = downloadPptx;
