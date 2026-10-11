/* Browser fixer for Hebrew Word files produced by ChatGPT.
   The file stays in the browser. Nothing is uploaded. */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.HebrewFix = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
  const XML = 'http://www.w3.org/XML/1998/namespace';

  const HEB = /[\u0590-\u05FF\uFB1D-\uFB4F]/;
  const LAT = /[A-Za-z\u00C0-\u024F]/;
  const BULLET_SYMS = new Set(['F0B7', 'F0A7', 'F0FC', 'F0D8', 'F076', 'F0A8', 'F0E8', '00B7', '2022']);
  const LEADING_BULLET = /^[\s\u00A0]*(?:[•●▪■◦‣∙·○]|[-–*](?=\s))[\s\u00A0]*/;
  const LEADING_SPACE = /^[ \u00A0]+/;
  const ARROWS = { '→': '←', '←': '→', '⇒': '⇐', '⇐': '⇒', '⟶': '⟵', '⟵': '⟶' };
  const ARROW_RE = /[→←⇒⇐⟶⟵]/;
  const REVERSED_DURATION = /^\s*(?<detail>.+?)\s+דקות\s*[–-]\s*(?<label>.+?)\s+(?<number>\d+)\s*$/;
  const REPEATED_DURATION = /^\s*(?<number>\d+)\s+\1\s*דקות(?![\u0590-\u05FF])/;
  const LONE_LETTER = /(?<![\u05D0-\u05EA\u05F3\u05F4'"])[\u05D0-\u05EA](?![\u05D0-\u05EA\u05F3\u05F4'"])/;
  const TRAILING_NUMBER = /^[\s\u00A0]*[\u05D0-\u05EA].*[\u05D0-\u05EA][\s\u00A0\-–]+\d+[\s\u00A0]*$/;
  const REPEATED_NUMBER = /(?<!\d)(\d+)[\s\u00A0]+\1(?!\d)/;
  const PART_RE = /^word\/(header\d*|footer\d*|footnotes|endnotes)\.xml$/;

  const PPR_ORDER = [
    'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
    'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens',
    'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN',
    'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing',
    'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
    'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange',
  ];
  const RPR_ORDER = [
    'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike',
    'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish',
    'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight',
    'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang',
    'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange',
  ];

  const REASONS = {
    lone: 'יש אות עברית בודדת. מילה נחתכה לשתיים, וכדאי לתקן אותה ידנית.',
    trailing: 'השורה נגמרת במספר. יכול להיות שהסדר שלה הפוך.',
    repeated: 'אותו מספר מופיע פעמיים ברצף.',
  };

  function local(el) {
    const name = el.localName || el.nodeName || '';
    const cut = name.indexOf(':');
    return cut === -1 ? name : name.slice(cut + 1);
  }

  function kids(el) {
    return [...el.childNodes].filter((node) => node.nodeType === 1);
  }

  function child(el, name) {
    return kids(el).find((node) => local(node) === name) || null;
  }

  function descendants(el, name) {
    const out = [];
    for (const node of kids(el)) {
      if (!name || local(node) === name) out.push(node);
      out.push(...descendants(node, name));
    }
    return out;
  }

  function attr(el, name) {
    if (!el) return null;
    const value = el.getAttributeNS ? el.getAttributeNS(W, name) : null;
    return value == null || value === '' ? el.getAttribute('w:' + name) || el.getAttribute(name) : value;
  }

  function setAttr(el, name, value) {
    el.setAttributeNS(W, 'w:' + name, value);
  }

  function makeEl(doc, name) {
    return doc.createElementNS(W, 'w:' + name);
  }

  function ensure(parent, name, order) {
    const found = child(parent, name);
    if (found) return found;
    const created = makeEl(parent.ownerDocument, name);
    const rank = order.indexOf(name);
    for (const node of kids(parent)) {
      const pos = order.indexOf(local(node));
      if (pos !== -1 && pos > rank) {
        parent.insertBefore(created, node);
        return created;
      }
    }
    parent.appendChild(created);
    return created;
  }

  function ensureFirst(parent, name) {
    const found = child(parent, name);
    if (found) return found;
    const created = makeEl(parent.ownerDocument, name);
    parent.insertBefore(created, parent.firstChild);
    return created;
  }

  function textOf(el, name) {
    return descendants(el, name).map((node) => node.textContent || '').join('');
  }

  function runText(run) {
    return textOf(run, 't');
  }

  function paraText(p) {
    return textOf(p, 't');
  }

  function isPlainBreak(el) {
    return local(el) === 'br' && (!attr(el, 'type') || attr(el, 'type') === 'textWrapping');
  }

  function count(re, text) {
    return (text.match(new RegExp(re.source, 'g')) || []).length;
  }

  function parse(xml) {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const root = doc.documentElement;
    if (!root || local(root) === 'parsererror') throw new Error('bad xml');
    return doc;
  }

  function dump(doc) {
    const xml = new XMLSerializer().serializeToString(doc.documentElement || doc);
    return xml.startsWith('<?xml')
      ? xml
      : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + xml;
  }

  function fixRpr(rpr) {
    if (!child(rpr, 'rtl')) ensure(rpr, 'rtl', RPR_ORDER);
    for (const [latin, cs] of [['b', 'bCs'], ['i', 'iCs']]) {
      const src = child(rpr, latin);
      const val = src ? attr(src, 'val') : null;
      if (src && !child(rpr, cs) && val !== '0' && val !== 'false') ensure(rpr, cs, RPR_ORDER);
    }
    const sz = child(rpr, 'sz');
    if (sz && !child(rpr, 'szCs') && attr(sz, 'val')) setAttr(ensure(rpr, 'szCs', RPR_ORDER), 'val', attr(sz, 'val'));
    const fonts = child(rpr, 'rFonts');
    if (fonts && !attr(fonts, 'cs') && !attr(fonts, 'cstheme') && attr(fonts, 'ascii')) {
      setAttr(fonts, 'cs', attr(fonts, 'ascii'));
    }
    setAttr(ensure(rpr, 'lang', RPR_ORDER), 'bidi', 'he-IL');
  }

  function fixRunProps(run) {
    const text = runText(run);
    if (child(run, 'instrText')) return;
    if (!text && !child(run, 'tab')) return;
    if (LAT.test(text) && !HEB.test(text)) return;
    fixRpr(ensureFirst(run, 'rPr'));
  }

  function isSimpleRun(run) {
    const body = kids(run).filter((node) => local(node) !== 'rPr');
    return body.length === 1 && local(body[0]) === 't';
  }

  function rprKey(run) {
    const rpr = child(run, 'rPr');
    return rpr ? new XMLSerializer().serializeToString(rpr) : '';
  }

  function mergeRuns(p) {
    let prev = null;
    for (const node of kids(p)) {
      if (local(node) === 'r' && isSimpleRun(node)) {
        if (prev && rprKey(prev) === rprKey(node)) {
          const target = child(prev, 't');
          target.textContent = (target.textContent || '') + (child(node, 't').textContent || '');
          target.setAttributeNS(XML, 'xml:space', 'preserve');
          p.removeChild(node);
          continue;
        }
        prev = node;
      } else if (!['bookmarkStart', 'bookmarkEnd', 'proofErr'].includes(local(node))) {
        prev = null;
      }
    }
  }

  function addAfter(anchor, node) {
    anchor.parentNode.insertBefore(node, anchor.nextSibling);
  }

  function splitAtBreaks(p) {
    const runsWithBr = new Set(kids(p).filter((node) => local(node) === 'r' && kids(node).some(isPlainBreak)));
    if (!runsWithBr.size) return [p];
    const ppr = child(p, 'pPr');
    const segments = [[]];
    for (const node of kids(p)) {
      if (node === ppr) continue;
      if (runsWithBr.has(node)) {
        const rpr = child(node, 'rPr');
        let cur = makeEl(p.ownerDocument, 'r');
        if (rpr) cur.appendChild(rpr.cloneNode(true));
        for (const part of kids(node)) {
          if (part === rpr) continue;
          if (isPlainBreak(part)) {
            if (kids(cur).length > (rpr ? 1 : 0)) segments[segments.length - 1].push(cur);
            segments.push([]);
            cur = makeEl(p.ownerDocument, 'r');
            if (rpr) cur.appendChild(rpr.cloneNode(true));
          } else {
            cur.appendChild(part.cloneNode(true));
          }
        }
        if (kids(cur).length > (rpr ? 1 : 0)) segments[segments.length - 1].push(cur);
      } else {
        segments[segments.length - 1].push(node.cloneNode(true));
      }
    }

    const segText = (seg) => seg.map((el) => textOf(el, 't')).join('').trim();
    const merged = [];
    for (const seg of segments) {
      if (segText(seg) || !merged.length) merged.push(seg);
      else merged[merged.length - 1].push(...seg);
    }
    if (merged.filter((seg) => segText(seg)).length < 2) return [p];

    const sect = ppr ? child(ppr, 'sectPr') : null;
    if (sect) ppr.removeChild(sect);
    for (const node of kids(p)) {
      if (node !== ppr) p.removeChild(node);
    }
    const result = [p];
    for (const node of merged[0]) p.appendChild(node);
    let anchor = p;
    for (const seg of merged.slice(1)) {
      const np = makeEl(p.ownerDocument, 'p');
      if (ppr) np.appendChild(ppr.cloneNode(true));
      for (const node of seg) np.appendChild(node);
      addAfter(anchor, np);
      anchor = np;
      result.push(np);
    }
    if (sect) ensureFirst(result[result.length - 1], 'pPr').appendChild(sect);
    return result;
  }

  function removeLeading(p, pattern) {
    const match = paraText(p).match(pattern);
    if (!match || !match[0]) return false;
    let left = match[0].length;
    for (const node of descendants(p, 't')) {
      if (left <= 0) break;
      const text = node.textContent || '';
      const cut = Math.min(left, text.length);
      node.textContent = text.slice(cut);
      left -= cut;
    }
    return true;
  }

  function dropSymbolBullet(p) {
    for (const run of descendants(p, 'r')) {
      const sym = child(run, 'sym');
      if (sym && BULLET_SYMS.has((attr(sym, 'char') || '').toUpperCase())) {
        run.removeChild(sym);
        return true;
      }
      if (runText(run).trim()) return false;
    }
    return false;
  }

  function replaceParagraphText(p, text) {
    const runs = kids(p).filter((node) => local(node) === 'r');
    const template = runs.find((run) => runText(run).trim()) || runs[0] || null;
    const replacement = makeEl(p.ownerDocument, 'r');
    const rpr = template ? child(template, 'rPr') : null;
    if (rpr) replacement.appendChild(rpr.cloneNode(true));
    const node = makeEl(p.ownerDocument, 't');
    node.textContent = text;
    replacement.appendChild(node);
    for (const run of runs) p.removeChild(run);
    p.appendChild(replacement);
  }

  function repairDurationText(p) {
    const text = paraText(p);
    const reversed = text.match(REVERSED_DURATION);
    if (reversed) {
      const detail = reversed.groups.detail.trim().replace(/^[,.:]+|[,.:]+$/g, '');
      const label = reversed.groups.label.trim().replace(/^[,.:]+|[,.:]+$/g, '');
      replaceParagraphText(p, `${reversed.groups.number} דקות – ${label}: ${detail}.`);
      return;
    }
    const repeated = text.match(REPEATED_DURATION);
    if (repeated) {
      replaceParagraphText(p, text.replace(REPEATED_DURATION, `${repeated.groups.number} דקות`).trim());
    }
  }

  function fixParagraph(p) {
    const text = paraText(p);
    LAT.lastIndex = 0;
    HEB.lastIndex = 0;
    if (LAT.test(text) && !HEB.test(text)) {
      LAT.lastIndex = 0;
      HEB.lastIndex = 0;
      return;
    }
    LAT.lastIndex = 0;
    HEB.lastIndex = 0;
    repairDurationText(p);
    const ppr = ensureFirst(p, 'pPr');
    const wasLtr = !child(ppr, 'bidi');
    if (wasLtr) {
      ensure(ppr, 'bidi', PPR_ORDER);
      const ind = child(ppr, 'ind');
      if (ind) ppr.removeChild(ind);
    }
    const jc = child(ppr, 'jc');
    if (jc && ['left', 'right', 'start', 'end'].includes(attr(jc, 'val'))) ppr.removeChild(jc);
    const mark = child(ppr, 'rPr');
    if (mark) fixRpr(mark);
    for (const node of descendants(p, 't')) {
      if (node.textContent && ARROW_RE.test(node.textContent)) {
        node.textContent = node.textContent.replace(new RegExp(ARROW_RE.source, 'g'), (arrow) => ARROWS[arrow]);
      }
    }
    for (const run of descendants(p, 'r')) fixRunProps(run);
    mergeRuns(p);
  }

  function bulletNumId(root) {
    if (root._bulletNumId) return root._bulletNumId;
    const abstracts = kids(root).filter((node) => local(node) === 'abstractNum');
    const nums = kids(root).filter((node) => local(node) === 'num');
    const absId = Math.max(-1, ...abstracts.map((node) => Number(attr(node, 'abstractNumId')) || 0)) + 1;
    const numId = Math.max(0, ...nums.map((node) => Number(attr(node, 'numId')) || 0)) + 1;
    const doc = root.ownerDocument;
    const abs = makeEl(doc, 'abstractNum');
    setAttr(abs, 'abstractNumId', String(absId));
    const nsid = makeEl(doc, 'nsid');
    setAttr(nsid, 'val', Math.floor(Math.random() * 0xffffffff).toString(16).padStart(8, '0').toUpperCase());
    abs.appendChild(nsid);
    const multi = makeEl(doc, 'multiLevelType');
    setAttr(multi, 'val', 'hybridMultilevel');
    abs.appendChild(multi);
    for (let level = 0; level < 3; level += 1) {
      const lvl = makeEl(doc, 'lvl');
      setAttr(lvl, 'ilvl', String(level));
      for (const [name, value] of [['start', '1'], ['numFmt', 'bullet'], ['lvlText', '•◦▪'[level]], ['lvlJc', 'left']]) {
        const node = makeEl(doc, name);
        setAttr(node, 'val', value);
        lvl.appendChild(node);
      }
      const ppr = makeEl(doc, 'pPr');
      const ind = makeEl(doc, 'ind');
      setAttr(ind, 'left', String(720 * (level + 1)));
      setAttr(ind, 'hanging', '360');
      ppr.appendChild(ind);
      lvl.appendChild(ppr);
      const rpr = makeEl(doc, 'rPr');
      const fonts = makeEl(doc, 'rFonts');
      setAttr(fonts, 'ascii', 'Arial');
      setAttr(fonts, 'hAnsi', 'Arial');
      setAttr(fonts, 'cs', 'Arial');
      setAttr(fonts, 'hint', 'default');
      rpr.appendChild(fonts);
      lvl.appendChild(rpr);
      abs.appendChild(lvl);
    }
    if (abstracts.length) addAfter(abstracts[abstracts.length - 1], abs);
    else if (nums.length) root.insertBefore(abs, nums[0]);
    else root.appendChild(abs);

    const num = makeEl(doc, 'num');
    setAttr(num, 'numId', String(numId));
    const link = makeEl(doc, 'abstractNumId');
    setAttr(link, 'val', String(absId));
    num.appendChild(link);
    const latestNums = kids(root).filter((node) => local(node) === 'num');
    if (latestNums.length) addAfter(latestNums[latestNums.length - 1], num);
    else root.appendChild(num);
    root._bulletNumId = numId;
    return numId;
  }

  function headingIds(stylesDoc) {
    const ids = new Set();
    if (!stylesDoc) return ids;
    for (const style of kids(stylesDoc.documentElement).filter((node) => local(node) === 'style')) {
      const name = child(style, 'name');
      const label = (name ? attr(name, 'val') : '').toLowerCase();
      const ppr = child(style, 'pPr');
      if (/^(heading|title|subtitle)/.test(label) || (ppr && child(ppr, 'outlineLvl'))) ids.add(attr(style, 'styleId'));
    }
    return ids;
  }

  function isHeading(p, ids) {
    const ppr = child(p, 'pPr');
    const style = ppr ? child(ppr, 'pStyle') : null;
    if (style && ids.has(attr(style, 'val'))) return true;
    return Boolean(ppr && child(ppr, 'outlineLvl'));
  }

  function startsLikeBullet(p) {
    if (LEADING_BULLET.test(paraText(p))) return true;
    for (const run of descendants(p, 'r')) {
      const sym = child(run, 'sym');
      if (sym && BULLET_SYMS.has((attr(sym, 'char') || '').toUpperCase())) return true;
      if (runText(run).trim()) return false;
    }
    return false;
  }

  function bodyGroups(body, ids) {
    const groups = [];
    let cur = [];
    for (const el of kids(body)) {
      const ppr = child(el, 'pPr');
      const isText = local(el) === 'p' && paraText(el).trim() && !isHeading(el, ids) && !(ppr && child(ppr, 'numPr'));
      if (isText) cur.push(el);
      else if (cur.length) {
        groups.push(cur);
        cur = [];
      }
    }
    if (cur.length) groups.push(cur);
    return groups;
  }

  function convertLists(body, ids, numbering) {
    for (const group of bodyGroups(body, ids)) {
      const bullets = group.filter(startsLikeBullet).length;
      const indented = group.filter((p) => LEADING_SPACE.test(paraText(p))).length;
      if (!(bullets || (group.length >= 2 && indented >= Math.ceil(group.length / 2)))) continue;
      const numId = bulletNumId(numbering);
      for (const p of group) {
        dropSymbolBullet(p);
        removeLeading(p, LEADING_BULLET);
        removeLeading(p, /^[\s\u00A0]+/);
        const ppr = ensureFirst(p, 'pPr');
        for (const name of ['ind', 'jc']) {
          const el = child(ppr, name);
          if (el) ppr.removeChild(el);
        }
        const numPr = ensure(ppr, 'numPr', PPR_ORDER);
        for (const node of kids(numPr)) numPr.removeChild(node);
        const ilvl = makeEl(p.ownerDocument, 'ilvl');
        setAttr(ilvl, 'val', '0');
        const id = makeEl(p.ownerDocument, 'numId');
        setAttr(id, 'val', String(numId));
        numPr.appendChild(ilvl);
        numPr.appendChild(id);
      }
    }
  }

  function fixStyles(doc) {
    const root = doc.documentElement;
    const defaults = ensureFirst(root, 'docDefaults');
    let rprDefault = child(defaults, 'rPrDefault');
    if (!rprDefault) {
      rprDefault = makeEl(doc, 'rPrDefault');
      defaults.insertBefore(rprDefault, defaults.firstChild);
    }
    setAttr(ensure(ensureFirst(rprDefault, 'rPr'), 'lang', RPR_ORDER), 'bidi', 'he-IL');
    let pprDefault = child(defaults, 'pPrDefault');
    if (!pprDefault) {
      pprDefault = makeEl(doc, 'pPrDefault');
      defaults.appendChild(pprDefault);
    }
    const ppr = ensureFirst(pprDefault, 'pPr');
    if (!child(ppr, 'bidi')) ensure(ppr, 'bidi', PPR_ORDER);
    for (const lang of descendants(root, 'lang')) {
      if (attr(lang, 'bidi') && attr(lang, 'bidi') !== 'he-IL') setAttr(lang, 'bidi', 'he-IL');
    }
    for (const style of kids(root).filter((node) => local(node) === 'style')) {
      if (attr(style, 'type') !== 'paragraph') continue;
      const sppr = child(style, 'pPr');
      const jc = sppr ? child(sppr, 'jc') : null;
      if (jc && ['left', 'right', 'start', 'end'].includes(attr(jc, 'val'))) sppr.removeChild(jc);
      const srpr = child(style, 'rPr');
      if (srpr && child(srpr, 'b') && !child(srpr, 'bCs')) ensure(srpr, 'bCs', RPR_ORDER);
    }
  }

  function collectWarnings(body) {
    const warnings = [];
    for (const p of descendants(body, 'p')) {
      const text = paraText(p).trim();
      if (!HEB.test(text)) continue;
      const reasons = [];
      if (LONE_LETTER.test(text)) reasons.push(REASONS.lone);
      if (TRAILING_NUMBER.test(text)) reasons.push(REASONS.trailing);
      if (REPEATED_NUMBER.test(text)) reasons.push(REASONS.repeated);
      if (reasons.length) warnings.push({ text: text.slice(0, 180), reasons });
    }
    return warnings;
  }

  function plainAttr(el, name) {
    return el.getAttribute(name) || el.getAttributeNS(CT, name) || el.getAttributeNS(PKG_REL, name);
  }

  function addNumberingPart(zip, files) {
    const types = parse(files['[Content_Types].xml']);
    const overrides = descendants(types.documentElement, 'Override');
    if (!overrides.some((node) => plainAttr(node, 'PartName') === '/word/numbering.xml')) {
      const override = types.createElementNS(CT, 'Override');
      override.setAttribute('PartName', '/word/numbering.xml');
      override.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml');
      types.documentElement.appendChild(override);
      files['[Content_Types].xml'] = dump(types);
    }
    const rels = parse(files['word/_rels/document.xml.rels']);
    const existing = descendants(rels.documentElement, 'Relationship');
    if (!existing.some((node) => plainAttr(node, 'Target') === 'numbering.xml')) {
      const ids = new Set(existing.map((node) => plainAttr(node, 'Id')));
      let n = 1;
      while (ids.has('rId' + n)) n += 1;
      const rel = rels.createElementNS(PKG_REL, 'Relationship');
      rel.setAttribute('Id', 'rId' + n);
      rel.setAttribute('Type', REL + '/numbering');
      rel.setAttribute('Target', 'numbering.xml');
      rels.documentElement.appendChild(rel);
      files['word/_rels/document.xml.rels'] = dump(rels);
    }
    for (const [name, xml] of Object.entries(files)) {
      if (name === '[Content_Types].xml' || name === 'word/_rels/document.xml.rels') zip.file(name, xml);
    }
  }

  async function fixDocx(data) {
    if (typeof JSZip === 'undefined') throw new Error('JSZip is missing');
    const zip = await JSZip.loadAsync(data);
    const files = {};
    const changed = new Set();
    const read = async (name) => (zip.files[name] && !zip.files[name].dir ? zip.files[name].async('string') : null);
    const save = (name, xml) => {
      files[name] = xml;
      changed.add(name);
    };
    files['word/document.xml'] = await read('word/document.xml');
    if (!files['word/document.xml']) {
      return { ok: false, error: 'הקובץ לא נראה כמו קובץ וורד תקין.' };
    }
    for (const name of ['word/styles.xml', 'word/numbering.xml', '[Content_Types].xml', 'word/_rels/document.xml.rels']) {
      files[name] = await read(name);
    }
    for (const name of Object.keys(zip.files)) {
      if (PART_RE.test(name)) files[name] = await read(name);
    }

    const doc = parse(files['word/document.xml']);
    const body = child(doc.documentElement, 'body');
    const all = paraText(body);
    if (count(HEB, all) <= count(LAT, all)) {
      return { ok: false, error: 'כמעט אין עברית בקובץ, אז לא שינינו אותו.' };
    }
    const styles = files['word/styles.xml'] ? parse(files['word/styles.xml']) : null;
    const ids = headingIds(styles);
    for (const p of [...descendants(body, 'p')]) splitAtBreaks(p);

    let numbering = null;
    if (files['word/numbering.xml']) numbering = parse(files['word/numbering.xml']);
    else numbering = parse('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>');
    convertLists(body, ids, numbering.documentElement);
    if (!numbering.documentElement._bulletNumId) numbering = null;

    for (const p of descendants(body, 'p')) {
      removeLeading(p, LEADING_SPACE);
      fixParagraph(p);
    }
    save('word/document.xml', dump(doc));

    for (const name of Object.keys(files)) {
      if (!PART_RE.test(name) || !files[name]) continue;
      const part = parse(files[name]);
      for (const p of descendants(part.documentElement, 'p')) fixParagraph(p);
      save(name, dump(part));
    }
    if (styles) {
      fixStyles(styles);
      save('word/styles.xml', dump(styles));
    }
    if (numbering) {
      const isNew = !zip.files['word/numbering.xml'];
      save('word/numbering.xml', dump(numbering));
      if (isNew) addNumberingPart(zip, files);
    }
    for (const name of changed) zip.file(name, files[name]);
    const warnings = collectWarnings(child(parse(files['word/document.xml']).documentElement, 'body'));
    const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    return { ok: true, bytes, warnings };
  }

  return { fixDocx };
});
