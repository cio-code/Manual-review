/**
 * JHCSC Manual Review - reads a Word (.docx) review and turns it into comments.
 * Runs in the browser; the file is never uploaded. Needs JSZip (js/vendor/jszip/jszip.min.js).
 *
 * It understands three kinds of review documents (and combines them if a file has more than one):
 *   1. A review memo: headings, then numbered or bulleted points, e.g. "Must fix before approval" / "Sec. 35: Make ... Fix: ..."
 *   2. Real Word comments (Review > New Comment), including the text they are attached to
 *   3. A table with a Comment column (and optionally Page/Section, Text, Suggested change columns)
 *
 * WordReview.parse(arrayBuffer) -> Promise<{ source, items: [{ location, quote, comment, suggestion, group }] }>
 */
(function () {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const LIMIT = { location: 200, quote: 2000, comment: 5000, suggestion: 2000, items: 200 };

  const squash = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
  const clip = (s, n) => { s = squash(s); return s.length > n ? s.slice(0, n - 1).trimEnd() + '\u2026' : s; };
  const kids = (el, name) => Array.from(el.childNodes).filter(n => n.nodeType === 1 && n.namespaceURI === W && n.localName === name);
  const attr = (el, name) => (el && (el.getAttributeNS(W, name) || el.getAttribute('w:' + name))) || '';

  /** Visible text of an element (deleted tracked text is not included). */
  function textOf(el) {
    let out = '';
    (function walk(n) {
      for (const c of n.childNodes) {
        if (c.nodeType !== 1) continue;
        if (c.namespaceURI === W && c.localName === 't') out += c.textContent;
        else if (c.namespaceURI === W && (c.localName === 'tab' || c.localName === 'br' || c.localName === 'cr')) out += ' ';
        else walk(c);
      }
    })(el);
    return squash(out);
  }

  function paraInfo(p, styleNames) {
    const ppr = kids(p, 'pPr')[0];
    let num = null, style = '';
    if (ppr) {
      const np = kids(ppr, 'numPr')[0];
      if (np) {
        const id = attr(kids(np, 'numId')[0], 'val'), lvl = attr(kids(np, 'ilvl')[0], 'val');
        if (id && id !== '0') num = { id: id, lvl: Number(lvl) || 0 };
      }
      style = attr(kids(ppr, 'pStyle')[0], 'val');
    }
    return { text: textOf(p), num: num, styleName: styleNames[style] || style };
  }

  /** A short line that is not a list item and does not end like a sentence is a heading (or the title). */
  function looksLikeHeading(info) {
    if (info.num || !info.text) return false;
    if (/^(title|subtitle|heading)/i.test(info.styleName)) return true;
    return info.text.length <= 90 && !/[.;,]$/.test(info.text);
  }

  const LOCATION = /^((?:Sec(?:tion)?s?\.?|Art(?:icle)?\.?|Page|Pages|pp?\.)\s*\d+[A-Za-z]?(?:\([A-Za-z0-9]+\))*)\s*(:)?\s*([\s\S]*)$/i;

  /** One memo point -> one comment. */
  function pointToComment(it) {
    let text = it.text, location = '', suggestion = '';
    const m = text.match(LOCATION);
    if (m) {
      location = m[1];
      if (m[2] && m[3]) text = m[3];                     // "Sec. 35: Make course-level..." -> the section becomes the location
    }
    if (!location) location = (it.heading || 'Review') + ' #' + it.n;
    const fx = text.match(/(?:^|\s)Fix:\s*([\s\S]+)$/);
    if (fx && fx.index > 0) { suggestion = fx[1]; text = text.slice(0, fx.index); }
    text = squash(text);
    return {
      location: clip(location, LIMIT.location),
      quote: '',
      comment: clip((it.heading ? '[' + it.heading + '] ' : '') + text, LIMIT.comment),
      suggestion: clip(suggestion, LIMIT.suggestion),
      group: it.heading || ''
    };
  }

  function tableToComments(tbl, heading) {
    const rows = kids(tbl, 'tr').map(tr => kids(tr, 'tc').map(textOf));
    if (rows.length < 2) return [];
    const head = rows[0].map(h => h.toLowerCase());
    const find = (re, not) => head.findIndex((h, i) => re.test(h) && not.indexOf(i) < 0);
    const c = find(/comment|issue|observ|remark|finding|concern|feedback/, []);
    if (c < 0) return [];
    const l = find(/page|section|\bsec\b|location|clause|article|\bref/, [c]);
    const s = find(/suggest|\bfix\b|recommend|revis|propos|action/, [c, l]);
    const q = find(/quote|text|passage|excerpt|provision/, [c, l, s]);
    const out = [];
    rows.slice(1).forEach((r, i) => {
      const comment = squash(r[c]);
      if (comment.length < 5) return;
      out.push({
        location: clip((l >= 0 && r[l]) || ((heading || 'Table') + ' row ' + (i + 1)), LIMIT.location),
        quote: q >= 0 ? clip(r[q], LIMIT.quote) : '',
        comment: clip(comment, LIMIT.comment),
        suggestion: s >= 0 ? clip(r[s], LIMIT.suggestion) : '',
        group: heading || 'Table'
      });
    });
    return out;
  }

  /** Real Word comments, with the text each one is attached to. */
  function wordComments(docDom, commentsDom, styleNames) {
    const byId = {};
    Array.from(commentsDom.getElementsByTagNameNS(W, 'comment')).forEach(c => {
      byId[attr(c, 'id')] = { author: attr(c, 'author') || 'Reviewer', text: kids(c, 'p').map(textOf).filter(Boolean).join(' ') };
    });
    const open = {}, anchor = {}, heading = {};
    let currentHeading = '';
    (function walk(node) {
      for (const c of node.childNodes) {
        if (c.nodeType !== 1 || c.namespaceURI !== W) { if (c.nodeType === 1) walk(c); continue; }
        if (c.localName === 'p') {
          const info = paraInfo(c, styleNames);
          if (looksLikeHeading(info)) currentHeading = info.text;
        }
        if (c.localName === 'commentRangeStart') { const id = attr(c, 'id'); open[id] = ''; heading[id] = currentHeading; }
        else if (c.localName === 'commentRangeEnd') { const id = attr(c, 'id'); if (id in open) { anchor[id] = open[id]; delete open[id]; } }
        else if (c.localName === 't') { for (const id in open) open[id] += c.textContent; }
        else walk(c);
      }
    })(docDom.documentElement);
    const out = [];
    Object.keys(byId).forEach(id => {
      const x = byId[id];
      if (squash(x.text).length < 5) return;
      const quote = squash(anchor[id] || '');
      const sec = quote.match(LOCATION);
      out.push({
        location: clip(sec ? sec[1] : (heading[id] ? 'Near: ' + heading[id] : 'Word comment'), LIMIT.location),
        quote: clip(quote, 600),
        comment: clip('[Word comment by ' + x.author + '] ' + x.text, LIMIT.comment),
        suggestion: '',
        group: 'Word comments'
      });
    });
    return out;
  }

  async function parse(buffer) {
    if (typeof JSZip === 'undefined') throw new Error('The Word reader did not load (js/vendor/jszip/jszip.min.js is missing).');
    let zip;
    try { zip = await JSZip.loadAsync(buffer); } catch (e) { throw new Error('That file is not a Word (.docx) file. Save it as .docx and try again.'); }
    const docFile = zip.file('word/document.xml');
    if (!docFile) throw new Error('That file is not a Word (.docx) file. Save it as .docx and try again.');
    const dom = text => {
      const d = new DOMParser().parseFromString(text, 'application/xml');
      if (d.getElementsByTagName('parsererror').length) throw new Error('The Word file could not be read.');
      return d;
    };
    const docDom = dom(await docFile.async('string'));

    const styleNames = {};
    const stylesFile = zip.file('word/styles.xml');
    if (stylesFile) {
      Array.from(dom(await stylesFile.async('string')).getElementsByTagNameNS(W, 'style')).forEach(s => {
        const n = kids(s, 'name')[0]; styleNames[attr(s, 'styleId')] = attr(n, 'val');
      });
    }

    const body = docDom.getElementsByTagNameNS(W, 'body')[0];
    const blocks = body ? Array.from(body.childNodes).filter(n => n.nodeType === 1 && n.namespaceURI === W) : [];
    const points = [], plain = [], tables = [], counter = {};
    let source = '', heading = '', pending = [], last = null, sawList = false;

    const settleHeadings = () => {                       // the last short line before the content is the heading; an earlier one is the title
      if (!pending.length) return;
      if (!source && pending.length > 1) source = pending[0];
      heading = pending[pending.length - 1];
      pending = [];
    };

    blocks.forEach(b => {
      if (b.localName === 'tbl') { settleHeadings(); tables.push.apply(tables, tableToComments(b, heading)); return; }
      if (b.localName !== 'p') return;
      const info = paraInfo(b, styleNames);
      if (!info.text) return;
      if (info.num) {
        settleHeadings(); sawList = true;
        if (info.num.lvl > 0 && last) { last.text += ' ' + info.text; return; }       // a sub-point belongs to its main point
        counter[heading] = (counter[heading] || 0) + 1;
        last = { heading: heading, text: info.text, n: counter[heading] };
        points.push(last);
      } else if (looksLikeHeading(info)) {
        pending.push(info.text); last = null;
      } else {
        settleHeadings();
        counter[heading] = (counter[heading] || 0) + 1;
        plain.push({ heading: heading, text: info.text, n: counter[heading] });
        last = null;
      }
    });
    if (!source && pending.length === 1 && !points.length && !plain.length && !tables.length) source = pending[0];

    const cf = zip.file('word/comments.xml');
    let items = points.map(pointToComment);
    // A file with no lists, no table and no Word comments is read as plain notes: every full paragraph is a comment.
    // If it has Word comments or a table, its paragraphs are the manual's own text and are NOT imported.
    if (!items.length && !cf && !tables.length) items = plain.filter(p => p.text.length >= 20).map(pointToComment);
    items = items.concat(tables);
    if (cf) items = items.concat(wordComments(docDom, dom(await cf.async('string')), styleNames));

    items = items.filter(x => x.comment.length >= 5);
    if (items.length > LIMIT.items) throw new Error('This file has ' + items.length + ' comments. Import up to ' + LIMIT.items + ' at a time (split the document).');
    return { source: source, items: items };
  }

  window.WordReview = { parse: parse };
})();
