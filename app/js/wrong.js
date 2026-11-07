/* 错词本 / 错题本：
 *  - 错词：默写答错自动记录，可单独重练
 *  - 错题：交卷时把答错的题记下来，可回到那一套重做
 */
(function () {
  const CET6 = (window.CET6 = window.CET6 || {});
  const { esc } = CET6.util;

  const TYPE_LABEL = { listening: '听力', cloze: '选词', matching: '长篇', reading: '仔细' };
  /** 听力是独立模块，重做时走它自己的入口 */
  const reopen = (type, paperId) => (type === 'listening' ? CET6.Listening.start(paperId) : CET6.Quiz.start(type, paperId));

  const W = (CET6.Wrong = {
    /* 错词本（按词库过滤） */
    words(size) {
      const all = CET6.Store.wrongWords();
      const list = size ? all.filter((w) => (w.lib || 6500) === size) : all;

      const items = list
        .map(
          (x) => `
        <div class="nb-item">
          <div class="w">${esc(x.w)}</div>
          <div class="pill">${x.ph ? '/' + esc(x.ph) + '/' : ''} 错 ${x.times} 次</div>
          <div class="d">${esc(x.def)}</div>
          <div class="foot">
            <button class="btn btn-ghost btn-sm" data-say="${esc(x.w)}">🔊</button>
            <button class="btn btn-ghost btn-sm" data-del="${esc(x.w)}">移出</button>
          </div>
        </div>`
        )
        .join('');

      CET6.UI.set(`
        <div class="wrap">
          <div class="section-title">
            <h2>错词本${size ? `（${size === 2500 ? '核心 2500' : '大纲 6500'}）` : ''}</h2>
            <span class="pill">${list.length} 个词</span>
            <span class="spacer"></span>
            <button class="btn btn-primary btn-sm" id="w-review" ${list.length ? '' : 'disabled'}>默写重练</button>
            <button class="btn btn-sm" id="w-clear" ${list.length ? '' : 'disabled'}>清空</button>
          </div>
          ${
            list.length
              ? `<div class="nb-list">${items}</div>`
              : CET6.UI.empty('✍️', '还没有默写错词', '在「默写」模式里打错的词会自动收进来')
          }
        </div>`);
      CET6.UI.backTo(() => CET6.App.home());

      document.querySelectorAll('[data-say]').forEach((b) => {
        b.onclick = () => CET6.util.speak(b.dataset.say);
      });
      document.querySelectorAll('[data-del]').forEach((b) => {
        b.onclick = () => {
          CET6.Store.clearWrongWord(b.dataset.del);
          W.words(size);
        };
      });
      const rv = document.getElementById('w-review');
      if (rv)
        rv.onclick = () => {
          const words = CET6.util.shuffle(list).map((x) => ({ w: x.w, ph: x.ph, def: x.def, stars: 0 }));
          CET6.Vocab.practiceWords(words, '错词重练', () => W.words(size));
        };
      const cl = document.getElementById('w-clear');
      if (cl)
        cl.onclick = async () => {
          const ok = await CET6.UI.confirm({ title: '清空错词', body: '只清空错词记录，不影响已学进度。', ok: '清空', danger: true });
          if (!ok) return;
          for (const x of list) CET6.Store.clearWrongWord(x.w);
          W.words(size);
        };
    },

    /* 错题本（按题型分组） */
    async render() {
      CET6.UI.loading();
      const papers = await CET6.Data.papersMeta();
      const book = CET6.Store.state.wrongBook;
      const byType = {};
      for (const it of book) (byType[it.type] = byType[it.type] || []).push(it);

      const groups = Object.entries(byType)
        .map(([type, items]) => {
          const paperIds = [...new Set(items.map((x) => x.paperId))];
          return `
            <div class="panel panel-pad" style="margin-bottom:16px">
              <div class="section-title">
                <h2 style="font-size:15px">${TYPE_LABEL[type] || type}</h2>
                <span class="pill">${items.length} 道 · 涉及 ${paperIds.length} 套</span>
              </div>
              ${paperIds
                .map((pid) => {
                  const meta = papers.find((p) => p.id === pid);
                  const list = items.filter((x) => x.paperId === pid).sort((a, b) => a.n - b.n);
                  return `<div class="wrong-row">
                    <div class="wrong-head">
                      <b>${esc(meta?.title || pid)}</b>
                      <span class="pill">第 ${list.map((x) => x.n).join('、')} 题</span>
                      <span class="spacer"></span>
                      <button class="btn btn-sm" data-redo="${pid}" data-type="${type}">重做这套</button>
                      <button class="btn btn-ghost btn-sm" data-drop="${pid}" data-type="${type}">标为已掌握</button>
                    </div>
                    ${list
                      .slice(0, 3)
                      .map(
                        (x) => `<div class="wrong-line"><span class="tag">${x.n}</span>
                          <span class="wrong-stem">${esc((x.stem || '').slice(0, 90))}${(x.stem || '').length > 90 ? '…' : ''}</span>
                          <span class="muted">错 ${x.times} 次</span></div>`
                      )
                      .join('')}
                    ${list.length > 3 ? `<div class="muted" style="margin-top:6px">其余 ${list.length - 3} 道…</div>` : ''}
                  </div>`;
                })
                .join('')}
            </div>`;
        })
        .join('');

      CET6.UI.set(`
        <div class="wrap">
          <div class="section-title">
            <h2>错题本</h2>
            <span class="pill">${book.length} 道待巩固</span>
            <span class="spacer"></span>
            <button class="btn btn-ghost btn-sm" id="wb-clear" ${book.length ? '' : 'disabled'}>清空</button>
          </div>
          ${
            book.length
              ? groups
              : CET6.UI.empty('📕', '还没有错题', '交卷时答错的题会自动记在这里，答对了会自动销账')
          }
        </div>`);
      CET6.UI.backTo(() => CET6.App.home());

      document.querySelectorAll('[data-redo]').forEach((b) => {
        b.onclick = async () => {
          const ok = await CET6.UI.confirm({
            title: '重做这套',
            body: '会清掉这套的作答记录并重新开始（错题本里的记录保留，答对后自动移除）。',
            ok: '开始重做',
          });
          if (!ok) return;
          delete CET6.Store.state.quiz[`${b.dataset.redo}:${b.dataset.type}`];
          CET6.Store.save();
          reopen(b.dataset.type, b.dataset.redo);
        };
      });
      document.querySelectorAll('[data-drop]').forEach((b) => {
        b.onclick = () => {
          const pid = b.dataset.drop;
          const type = b.dataset.type;
          CET6.Store.state.wrongBook = CET6.Store.state.wrongBook.filter((x) => !(x.paperId === pid && x.type === type));
          CET6.Store.save();
          W.render();
        };
      });
      const cl = document.getElementById('wb-clear');
      if (cl)
        cl.onclick = async () => {
          const ok = await CET6.UI.confirm({ title: '清空错题本', body: '只是清空清单，不影响作答记录。', ok: '清空', danger: true });
          if (!ok) return;
          CET6.Store.clearWrongBook();
          W.render();
        };
    },
  });
})();
