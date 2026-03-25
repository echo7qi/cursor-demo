/* 祈愿单项目复盘 · 动态看板：解析整体数据监测 CSV，按专题展示、搜索、本周/近7日新上线 */

(function () {
  const $ = (id) => document.getElementById(id);

  function esc(s) {
    if (s == null || s === '') return '';
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function toNum(v) {
    if (v == null || v === '') return null;
    const n = parseFloat(String(v).replace(/,/g, ''));
    return Number.isFinite(n) ? n : null;
  }

  /** 单元格字符串（去首尾空白，避免导出表含不可见空格导致匹配失败） */
  function val(r, key) {
    if (!r) return '';
    const v = r[key];
    if (v == null) return '';
    return String(v).trim();
  }

  /** 去掉表头 BOM、列名首尾空白，避免首列识别不到「数据分类」 */
  function normalizeRows(rows) {
    if (!rows || !rows.length) return rows || [];
    return rows.map((r) => {
      const out = {};
      Object.keys(r).forEach((k) => {
        const nk = k.replace(/^\uFEFF/, '').trim();
        out[nk] = r[k];
      });
      return out;
    });
  }

  /** 多文件合并：同一活动+数据分类+周期+用户类型 只保留一行（后读入的文件覆盖先读的，时间序由 wish-review.js 从旧到新） */
  function rowDedupeKey(r) {
    return [
      val(r, '活动标识'),
      val(r, '数据分类'),
      val(r, '数据周期'),
      val(r, '是否目标用户'),
    ].join('\x01');
  }

  function dedupeMonitoringRows(rows) {
    const m = new Map();
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      m.set(rowDedupeKey(r), r);
    }
    return Array.from(m.values());
  }

  function isSummaryAllRow(r) {
    return (
      val(r, '数据分类') === '汇总' &&
      val(r, '数据周期') === '汇总' &&
      val(r, '是否目标用户') === '全部'
    );
  }

  function periodNum(row) {
    const n = parseInt(String(val(row, '第x次祈愿') || '0'), 10);
    return Number.isFinite(n) ? n : 0;
  }

  function parseLaunchDate(raw) {
    if (raw == null) return null;
    const s = String(raw).trim().slice(0, 10);
    if (!s) return null;
    const t = Date.parse(s.replace(/\//g, '-'));
    if (Number.isNaN(t)) return null;
    return t;
  }

  function startOfWeekMondayMs() {
    const d = new Date();
    const day = d.getDay();
    const diff = day === 0 ? -6 : 1 - day;
    d.setDate(d.getDate() + diff);
    d.setHours(0, 0, 0, 0);
    return d.getTime();
  }

  function snapAll(rows, aid, n) {
    const key = `上线${n}日内`;
    const aidNorm = String(aid || '').trim();
    let allRow = null;
    let yesRow = null;
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (val(r, '活动标识') !== aidNorm || val(r, '数据分类') !== '当前累计' || val(r, '数据周期') !== key) {
        continue;
      }
      if (val(r, '是否目标用户') === '全部') allRow = r;
      if (val(r, '是否目标用户') === '是') yesRow = r;
    }
    return { pa: allRow, pyes: yesRow, n };
  }

  function snapRowN(rows, aid, maxDays) {
    let md = parseInt(String(maxDays), 10);
    if (Number.isNaN(md)) md = 9;
    const n = Math.min(9, Math.max(1, md));
    return snapAll(rows, aid, n);
  }

  /**
   * 与脚本一致：只认「汇总 / 汇总 / 全部」汇总行；同一专题下同一活动标识只保留一行（取期次更大者），
   * 避免导出重复行把「期数」和「专题数」撑大。
   */
  function buildTopicModels(rows) {
    const summaries = rows.filter(isSummaryAllRow);
    const rawSummaryCount = summaries.length;

    /** 专题名称 -> 活动标识 -> 行 */
    const byTopic = new Map();
    for (let i = 0; i < summaries.length; i++) {
      const r = summaries[i];
      const name = val(r, '专题名称');
      if (!name) continue;
      const aid = val(r, '活动标识');
      if (!aid) continue;
      if (!byTopic.has(name)) byTopic.set(name, new Map());
      const m = byTopic.get(name);
      const prev = m.get(aid);
      if (!prev || periodNum(r) >= periodNum(prev)) {
        m.set(aid, r);
      }
    }

    const topics = [];
    const weekStart = startOfWeekMondayMs();
    const now = Date.now();
    const sevenAgo = now - 7 * 86400000;

    let activityDedupCount = 0;
    byTopic.forEach((aidMap, name) => {
      const arr = Array.from(aidMap.values());
      activityDedupCount += arr.length;
      arr.sort((a, b) => periodNum(b) - periodNum(a));
      const latest = arr[0];
      const launchMs = parseLaunchDate(val(latest, '上线日期'));
      const inThisWeek = launchMs != null && launchMs >= weekStart;
      const in7d = launchMs != null && launchMs >= sevenAgo;
      topics.push({
        name,
        periods: arr,
        latest,
        launchMs,
        inThisWeek,
        in7d,
      });
    });

    topics.sort((a, b) => (b.launchMs || 0) - (a.launchMs || 0));
    return {
      topics,
      rawSummaryCount,
      activityDedupCount,
      topicCount: topics.length,
    };
  }

  let state = {
    rows: [],
    topics: [],
    fileName: '',
    selected: null,
    filter: '',
  };

  function fmtPct(x) {
    if (x == null) return '—';
    return `${(x * 100).toFixed(1)}%`;
  }

  function fmtInt(x) {
    if (x == null) return '—';
    return Math.round(x).toLocaleString('zh-CN');
  }

  function renderDetail(topicName) {
    const host = $('wishReviewDetailInner');
    if (!host) return;
    if (!topicName) {
      host.innerHTML =
        '<p class="muted wishReviewDash__empty">绑定后在左侧选择专题；支持按专题名称搜索。数据来自整体数据监测 CSV。</p>';
      return;
    }
    const t = state.topics.find((x) => x.name === topicName);
    if (!t) {
      host.innerHTML = '<p class="muted wishReviewDash__empty">未找到该专题</p>';
      return;
    }
    const { latest, periods } = t;
    const aid = val(latest, '活动标识');
    const days = toNum(val(latest, '已上线天数')) ?? 9;
    const { pa, n } = snapRowN(state.rows, aid, days);

    let kpiBlock = '';
    if (pa) {
      const rev = toNum(val(pa, '总收入'));
      const join = toNum(val(pa, '参与付费率'));
      const tgt = toNum(val(pa, '目标触达率'));
      const tuv = toNum(val(pa, '触达用户数'));
      kpiBlock = `
        <div class="wishReviewDash__kpis">
          <div class="wishReviewDash__kpi"><span class="wishReviewDash__kpiL">上线${n}日·累计收入</span><strong>${fmtInt(rev)}</strong></div>
          <div class="wishReviewDash__kpi"><span class="wishReviewDash__kpiL">参与付费率</span><strong>${join != null ? fmtPct(join) : '—'}</strong></div>
          <div class="wishReviewDash__kpi"><span class="wishReviewDash__kpiL">目标触达率</span><strong>${tgt != null ? fmtPct(tgt) : '—'}</strong></div>
          <div class="wishReviewDash__kpi"><span class="wishReviewDash__kpiL">触达 UV</span><strong>${fmtInt(tuv)}</strong></div>
        </div>`;
    } else {
      kpiBlock =
        '<p class="muted">未匹配到「当前累计·上线' +
        esc(Math.min(9, Math.max(1, Math.floor(days)))) +
        '日内·全部」行，可能导出口径不同。</p>';
    }

    const periodRows = periods
      .map((r) => {
        const p = periodNum(r);
        const pool = esc(val(r, '活动名称【修正】') || val(r, '活动名称') || '—');
        const ld = esc(val(r, '上线日期') || '—');
        const d0 = esc(val(r, '已上线天数') || '—');
        return `<tr><td>第 ${p} 期</td><td>${pool}</td><td>${ld}</td><td>${d0}</td></tr>`;
      })
      .join('');

    host.innerHTML = `
      <div class="wishReviewDash__detailHead">
        <h2 class="wishReviewDash__detailTitle">${esc(t.name)}</h2>
        <p class="muted wishReviewDash__detailMeta">品类：${esc(val(latest, '品类') || '—')} · 监测表最新期：第 ${periodNum(latest)} 期 · 上线 ${esc(val(latest, '上线日期') || '—')}</p>
      </div>
      ${kpiBlock}
      <h3 class="wishReviewDash__subTitle">专题内各期（汇总行）</h3>
      <div class="tableWrap">
        <table class="table table--compact">
          <thead><tr><th>期次</th><th>活动名称</th><th>上线日期</th><th>已上线天数</th></tr></thead>
          <tbody>${periodRows}</tbody>
        </table>
      </div>
      <p class="muted wishReviewDash__footNote">指标快照与脚本「上线满9日·当前累计」对齐思路一致；完整五维卡片请本地运行 <code>生成_人鱼全期结论表.py --topic</code> 生成 HTML。</p>
    `;
  }

  function topicMatchesFilter(t, q) {
    if (!q) return true;
    return t.name.toLowerCase().includes(q);
  }

  function renderList() {
    const q = state.filter.trim().toLowerCase();
    const filtered = state.topics.filter((t) => topicMatchesFilter(t, q));

    if (
      state.selected &&
      !filtered.some((x) => x.name === state.selected)
    ) {
      state.selected = filtered[0] ? filtered[0].name : null;
      renderDetail(state.selected);
    }

    const weekList = filtered.filter((t) => t.inThisWeek);
    const recent7 = filtered.filter((t) => t.in7d && !t.inThisWeek);
    const rest = filtered.filter((t) => !t.inThisWeek && !t.in7d);

    function itemRow(t, badge) {
      const active = state.selected === t.name ? ' is-active' : '';
      const b = badge
        ? `<span class="wishReviewDash__badge">${esc(badge)}</span>`
        : '';
      const ld = t.launchMs
        ? new Date(t.launchMs).toLocaleDateString('zh-CN')
        : '—';
      return `<button type="button" class="wishReviewDash__topicBtn${active}" data-topic="${esc(t.name)}">
        <span class="wishReviewDash__topicName">${esc(t.name)}${b}</span>
        <span class="wishReviewDash__topicSub">第 ${periodNum(t.latest)} 期 · ${esc(ld)}</span>
      </button>`;
    }

    const elWeek = $('wishReviewWeekList');
    const el7 = $('wishReview7dList');
    const elAll = $('wishReviewAllList');
    if (elWeek) {
      elWeek.innerHTML = weekList.length
        ? weekList.map((t) => itemRow(t, '本周')).join('')
        : '<p class="muted wishReviewDash__empty">暂无</p>';
    }
    if (el7) {
      el7.innerHTML = recent7.length
        ? recent7.map((t) => itemRow(t, '近7日')).join('')
        : '<p class="muted wishReviewDash__empty">暂无（已在「本周」列出除外）</p>';
    }
    if (elAll) {
      const show = [...weekList, ...recent7, ...rest];
      elAll.innerHTML = show.length
        ? show.map((t) => itemRow(t, '')).join('')
        : '<p class="muted wishReviewDash__empty">无匹配专题</p>';
    }
  }

  function setSelected(name) {
    state.selected = name;
    renderList();
    renderDetail(name);
  }

  async function loadFromBinding() {
    const status = $('wishReviewDashStatus');
    const src = $('wishReviewDashSource');
    if (!window.wishReviewReadMonitorCsv) {
      if (status) status.textContent = '脚本未就绪。';
      return;
    }
    if (status) status.textContent = '正在读取监测表…';
    const r = await window.wishReviewReadMonitorCsv();
    if (!r.ok) {
      if (status) status.textContent = r.error || '读取失败';
      if (src) src.textContent = '';
      state.rows = [];
      state.topics = [];
      renderList();
      renderDetail(null);
      return;
    }
    const parts = r.parts && r.parts.length ? r.parts : null;
    let mergedRows = [];
    try {
      if (parts) {
        for (let pi = 0; pi < parts.length; pi++) {
          const parsed = Papa.parse(parts[pi].text, {
            header: true,
            skipEmptyLines: 'greedy',
          });
          if (parsed.errors && parsed.errors.length) {
            console.warn('[wish-review-dynamic]', parts[pi].name, parsed.errors);
          }
          mergedRows = mergedRows.concat(normalizeRows(parsed.data || []));
        }
      } else if (r.text) {
        const parsed = Papa.parse(r.text, {
          header: true,
          skipEmptyLines: 'greedy',
        });
        if (parsed.errors && parsed.errors.length) {
          console.warn('[wish-review-dynamic]', parsed.errors);
        }
        mergedRows = normalizeRows(parsed.data || []);
      } else {
        if (status) {
          status.textContent = '读取结果缺少 CSV 内容，请刷新页面后重试。';
        }
        return;
      }
    } catch (e) {
      if (status) status.textContent = 'CSV 解析失败';
      return;
    }
    state.rows = dedupeMonitoringRows(mergedRows);
    const built = buildTopicModels(state.rows);
    state.topics = built.topics;
    state._meta = built;
    state.fileName = r.fileName;
    state.fileNames =
      r.fileNames ||
      (parts ? parts.map((p) => p.name) : r.fileName ? [r.fileName] : []);
    const mtime = new Date(r.lastModified).toLocaleString('zh-CN', { hour12: false });
    const nFiles = parts ? parts.length : 1;
    if (status) {
      status.textContent =
        (nFiles > 1
          ? `已合并 ${nFiles} 个监测 CSV → ${state.rows.length} 行（跨文件去重后）`
          : `已加载 ${state.rows.length} 行`) +
        ` · ${built.topicCount} 个专题 · ${built.activityDedupCount} 个祈愿活动` +
        `（汇总·全部 原始 ${built.rawSummaryCount} 行，专题内按活动标识去重）`;
    }
    if (src) {
      const label =
        r.fileNames && r.fileNames.length
          ? r.fileNames.join(' + ')
          : r.fileName || (parts && parts.map((p) => p.name).join(' + ')) || '—';
      src.textContent = `${label} · ${mtime}`;
    }

    if (!state.selected || !state.topics.some((x) => x.name === state.selected)) {
      state.selected = state.topics[0] ? state.topics[0].name : null;
    }
    renderList();
    renderDetail(state.selected);
  }

  function init() {
    const search = $('wishReviewTopicSearch');
    const reload = $('wishReviewReloadBtn');
    if (search) {
      search.addEventListener('input', () => {
        state.filter = search.value;
        renderList();
      });
    }
    if (reload) {
      reload.addEventListener('click', () => loadFromBinding());
    }
    const lists = ['wishReviewWeekList', 'wishReview7dList', 'wishReviewAllList'];
    lists.forEach((id) => {
      const el = $(id);
      if (!el) return;
      el.addEventListener('click', (ev) => {
        const btn = ev.target.closest('[data-topic]');
        if (!btn) return;
        setSelected(btn.getAttribute('data-topic'));
      });
    });

    document.addEventListener('wishreview:datasource-updated', () => loadFromBinding());

    loadFromBinding();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
