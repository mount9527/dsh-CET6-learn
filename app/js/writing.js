/* 作文书写：随机抽一道真题作文题，写 + 词数/句数统计 + 30 分钟倒计时。
 *
 * 题库不收录范文，这里也不生成：写完可以交给老师批改，或隔天自己按评分要点自查。
 */
(function () {
  const CET6 = (window.CET6 = window.CET6 || {});
  const { esc } = CET6.util;
  const LIMIT = 30 * 60 * 1000;
  const TARGET = [150, 200];

  const W = (CET6.Writing = {
    async start() {
      CET6.UI.loading('抽题中…');
      const papers = await CET6.Data.papersMeta();
      const pool = papers.filter((p) => p.hasWriting);
      if (!pool.length) {
        CET6.UI.set(`<div class="wrap">${CET6.UI.empty('📝', '题库里还没有作文题')}</div>`);
        return;
      }
      const last = CET6.Store.state.writing.__last;
      const fresh = pool.filter((p) => p.id !== last);
      const meta = CET6.util.pick(fresh.length ? fresh : pool);
      CET6.Store.state.writing.__last = meta.id;
      CET6.Store.save();
      const paper = await CET6.Data.paper(meta.id);
      const prompt = paper.writing?.prompt || '';
      const key = `${paper.id}:essay`;
      const draft = CET6.Store.state.writing[key] || { text: '' };
      CET6.Store.state.writing[key] = draft;

      CET6.UI.set(`
        <div class="wrap">
          <div class="section-title">
            <h2>作文书写</h2>
            <span class="pill">${esc(paper.title)}</span>
            <span class="spacer"></span>
            <span class="timer" id="w-timer" title="考试建议 30 分钟">30:00</span>
            <button class="btn btn-ghost btn-sm" id="w-redraw">换一题</button>
          </div>
          <div class="panel panel-pad" style="margin-bottom:16px">
            <div class="section-title"><h2 style="font-size:15px">题目</h2><span class="pill">建议 ${TARGET[0]}–${TARGET[1]} 词</span></div>
            <div class="cn-text" style="font-size:15.5px">${esc(prompt || '（本题题干解析失败，请参考真题原文）')}</div>
          </div>
          <div class="panel panel-pad">
            <div class="section-title"><h2 style="font-size:15px">我的作文</h2><span class="spacer"></span>
              <span class="pill" id="w-count">0 词 · 0 句</span>
            </div>
            <textarea class="text" id="w-input" style="min-height:340px" placeholder="First sentence…">${esc(draft.text || '')}</textarea>
            <div class="word-bar"><i id="w-bar"></i></div>
            <div class="row" style="margin-top:12px">
              <button class="btn btn-ghost btn-sm" id="w-clear">清空</button>
              <span class="muted" style="font-size:12.5px">内容会自动保存到本机</span>
            </div>
          </div>
        </div>`);
      CET6.UI.backTo(() => CET6.App.home());

      const input = document.getElementById('w-input');
      const countEl = document.getElementById('w-count');
      const bar = document.getElementById('w-bar');
      const update = () => {
        const t = input.value;
        const words = (t.trim().match(/[A-Za-z'’]+/g) || []).length;
        const sents = CET6.util.sentences(t).length;
        countEl.textContent = `${words} 词 · ${sents} 句`;
        const pct = Math.min(100, (words / TARGET[1]) * 100);
        bar.style.width = `${pct}%`;
        bar.className = words < TARGET[0] ? 'low' : words > TARGET[1] ? 'high' : 'ok';
      };
      update();
      input.oninput = () => {
        update();
        draft.text = input.value;
        CET6.Store.save();
      };

      /* 倒计时：切走暂停 */
      const timer = CET6.Timer.start({
        from: draft.elapsed || 0,
        limit: LIMIT,
        onTick: (ms, left) => {
          const el = document.getElementById('w-timer');
          if (!el) return;
          el.textContent = CET6.util.fmtTime(left || 0);
          el.classList.toggle('over', left === 0);
          draft.elapsed = ms;
        },
      });
      CET6.UI.onLeave(() => {
        timer.stop();
        draft.elapsed = timer.total;
        CET6.Store.save();
      });

      document.getElementById('w-clear').onclick = async () => {
        const ok = await CET6.UI.confirm({ title: '清空正文', body: '只清空这篇的草稿内容。', ok: '清空' });
        if (!ok) return;
        input.value = '';
        draft.text = '';
        update();
        CET6.Store.save();
      };
      document.getElementById('w-redraw').onclick = () => W.start();
    },
  });
})();
