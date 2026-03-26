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

  /** 单行：去掉表头 BOM、列名首尾空白 */
  function normalizeRowKeys(r) {
    if (!r || typeof r !== 'object') return {};
    const out = {};
    Object.keys(r).forEach((k) => {
      const nk = k.replace(/^\uFEFF/, '').trim();
      out[nk] = r[k];
    });
    return out;
  }

  function isSummaryAllRow(r) {
    return (
      val(r, '数据分类') === '汇总' &&
      val(r, '数据周期') === '汇总' &&
      val(r, '是否目标用户') === '全部'
    );
  }

  /**
   * 本页看板实际用到的行远少于全表：大量明细（分日、分渠道等）可丢弃，避免数据逐年膨胀时拖垮浏览器。
   * 保留：① 汇总/汇总/全部（建专题列表）；② 当前累计·上线1–9日内·全部或「是」（详情 KPI）。
   */
  function isLaunchWindowWithin9DaysRow(r) {
    if (val(r, '数据分类') !== '当前累计') return false;
    const dr = val(r, '数据周期');
    if (!/^上线\s*([1-9])\s*日内$/.test(dr)) return false;
    const ut = val(r, '是否目标用户');
    return ut === '全部' || ut === '是';
  }

  function isRowUsedByDashboard(r) {
    return isSummaryAllRow(r) || isLaunchWindowWithin9DaysRow(r);
  }

  /** 流式解析：只把「看板用行」推入数组，避免 Papa 一次性生成百万行 objects */
  function parseMonitoringCsvToSlimRows(text, logLabel) {
    const acc = [];
    const errs = [];
    Papa.parse(text, {
      header: true,
      skipEmptyLines: 'greedy',
      worker: false,
      error(e) {
        errs.push({ type: 'fatal', message: String((e && e.message) || e) });
      },
      step(results) {
        if (results.errors && results.errors.length) {
          for (let i = 0; i < results.errors.length; i++) errs.push(results.errors[i]);
        }
        const raw = results.data;
        if (raw == null || typeof raw !== 'object' || Array.isArray(raw)) return;
        const row = normalizeRowKeys(raw);
        if (!isRowUsedByDashboard(row)) return;
        acc.push(row);
      },
      complete(results) {
        if (results && results.errors && results.errors.length) {
          for (let i = 0; i < results.errors.length; i++) errs.push(results.errors[i]);
        }
      },
    });
    if (errs.length) {
      console.warn('[wish-review-dynamic]', logLabel || 'CSV', errs.slice(0, 10));
    }
    return acc;
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
    benchRows: [],
    layerRows: [],
    workRows: [],
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
    const readFn = window.wishReviewReadMonitorCsv;
    if (!readFn) {
      if (status) status.textContent = '脚本未就绪。';
      return;
    }
    if (status) {
      status.textContent = '正在解析整体数据监测（流式瘦身，仅保留看板用行）…';
    }
    const mainBlock = await readFn();
    if (!mainBlock.ok) {
      if (status) status.textContent = mainBlock.error || '读取失败';
      if (src) src.textContent = '';
      state.rows = [];
      state.benchRows = [];
      state.layerRows = [];
      state.workRows = [];
      state.topics = [];
      if (typeof window !== 'undefined') window.__WISH_REVIEW_BUNDLE_DATA__ = null;
      renderList();
      renderDetail(null);
      return;
    }
    const parts = mainBlock.parts && mainBlock.parts.length ? mainBlock.parts : null;
    let mergedRows = [];
    try {
      if (parts) {
        for (let pi = 0; pi < parts.length; pi++) {
          mergedRows = mergedRows.concat(
            parseMonitoringCsvToSlimRows(parts[pi].text, parts[pi].name),
          );
        }
      } else if (mainBlock.text) {
        mergedRows = parseMonitoringCsvToSlimRows(
          mainBlock.text,
          mainBlock.fileName || '监测表',
        );
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
    state.benchRows = [];
    state.layerRows = [];
    state.workRows = [];
    if (typeof window !== 'undefined') {
      window.__WISH_REVIEW_BUNDLE_DATA__ = {
        mainRowCount: state.rows.length,
        benchRowCount: 0,
        layerRowCount: 0,
        workRowCount: 0,
        loadedAt: Date.now(),
        note: '仅保留汇总行+当前累计·上线1–9日内；流式解析降低内存',
      };
    }
    const built = buildTopicModels(state.rows);
    state.topics = built.topics;
    state._meta = built;
    state.fileName = mainBlock.fileName;
    state.fileNames =
      mainBlock.fileNames ||
      (parts ? parts.map((p) => p.name) : mainBlock.fileName ? [mainBlock.fileName] : []);
    const mtime = new Date(mainBlock.lastModified).toLocaleString('zh-CN', { hour12: false });
    const nFiles = parts ? parts.length : 1;
    if (status) {
      status.textContent =
        (nFiles > 1
          ? `已合并 ${nFiles} 个监测 CSV → ${state.rows.length} 行（跨文件去重后）`
          : `已加载 ${state.rows.length} 行`) +
        ` · ${built.topicCount} 个专题 · ${built.activityDedupCount} 个祈愿活动` +
        `（汇总·全部 原始 ${built.rawSummaryCount} 行，专题内按活动标识去重）` +
        ' · 已瘦身解析（原表再大也只保留看板用行）；对标池等未载入，全量请用本地脚本';
    }
    if (src) {
      const label =
        mainBlock.fileNames && mainBlock.fileNames.length
          ? mainBlock.fileNames.join(' + ')
          : mainBlock.fileName || (parts && parts.map((p) => p.name).join(' + ')) || '—';
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
