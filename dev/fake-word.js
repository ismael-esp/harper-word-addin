// Dev only: fake `Office` and `Word` APIs for the task pane, backed by the document model in dev/index.html (window.parent.fakeDoc).
// It is lenient: real Word throws if you read a property you didn't load() and sync(); this mock doesn't.
(() => {
  const doc = window.parent.fakeDoc;
  const stats = doc.stats;

  window.Office = {
    context: { requirements: { isSetSupported: (set, v) => set === 'WordApi' && parseFloat(v) <= 1.7 } },
    onReady: cb => Promise.resolve().then(() => cb && cb({ host: 'Word' })),
  };

  function range(p, s, e) {
    return {
      load() {},
      get text() { return doc.paragraphs[p].slice(s, e); },
      select() { doc.select(p, s, e); },
      delete() { doc.edit(p, s, e, ''); },
      insertText(t, loc = 'Replace') {
        if (loc === 'Replace') doc.edit(p, s, e, t);
        else if (loc === 'Start' || loc === 'Before') doc.edit(p, s, s, t);
        else doc.edit(p, e, e, t);
      },
    };
  }

  function para(p) {
    return {
      load() {},
      get text() { return doc.paragraphs[p]; },
      search(q, opts = {}) {
        stats.search++;
        if (q.length > 255) throw new Error('SearchStringInvalidOrTooLong');
        const hay = opts.matchCase ? doc.paragraphs[p] : doc.paragraphs[p].toLowerCase();
        const needle = opts.matchCase ? q : q.toLowerCase();
        const items = [];
        for (let i = hay.indexOf(needle); needle && i !== -1; i = hay.indexOf(needle, i + needle.length)) items.push(range(p, i, i + q.length));
        return { items, load() {} };
      },
      getAnnotations() {
        stats.getAnnotations++;
        return { items: doc.marksOf(p).map(m => ({ id: m.id, delete() { doc.removeMark(m.id); } })), load() {} };
      },
      insertAnnotations(set) {
        stats.insertAnnotations++;
        return { value: set.critiques.map(c => doc.addMark(p, c)) };
      },
      insertText(t, loc = 'Replace') {
        if (loc === 'Replace') doc.edit(p, 0, doc.paragraphs[p].length, t);
        else if (loc === 'Start') doc.edit(p, 0, 0, t);
        else doc.edit(p, doc.paragraphs[p].length, doc.paragraphs[p].length, t);
      },
    };
  }

  window.Word = {
    run: async fn => {
      stats.run++;
      const ctx = {
        document: {
          body: { paragraphs: { load() {}, get items() { return doc.paragraphs.map((_, i) => para(i)); } } },
          onAnnotationClicked: { add(h) { doc.onAnnotationClicked = h; } },
        },
        sync: async () => { stats.sync++; await new Promise(r => setTimeout(r, 2)); }, // real sync is a round trip
      };
      return fn(ctx);
    },
  };
})();
