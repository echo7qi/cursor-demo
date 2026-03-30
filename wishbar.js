/* global Papa, Chart */

const $ = (id) => document.getElementById(id);

const DB_NAME = 'ops-dashboard-local-db';
const DB_STORE = 'kv';
const DB_KEY_DIR_HANDLE = 'boundDirHandle';

const COLS = {
  date: '日期',
  period: '统计周期',
  l1: '一级来源',
  l2: '二级来源',
  uv: '祈愿bar访问uv',
  drawUsers: '抽卡用户数',
  payUsers: '付费抽卡用户数',
  payAmount: '付费金额',
  activity: '活动名称【修正】',
};

const COLS_ALIASES = {
  date: ['日期', '周', 'date', 'week'],
  period: ['统计周期', '周期', 'period'],
  l1: ['一级来源', '一级', 'l1'],
  l2: ['二级来源', '二级', 'l2'],
  uv: ['祈愿bar访问uv', '祈愿bar访问UV', '访问uv', 'uv', '祈愿 bar 访问 uv'],
  drawUsers: ['抽卡用户数', '抽卡用户', 'drawUsers', 'draw_users'],
  payUsers: ['付费抽卡用户数', '付费用户数', 'payUsers', 'pay_users'],
  payAmount: ['付费金额', '付费总额', 'payAmount', 'pay_amount', '金额'],
  activity: ['活动名称【修正】', '活动名称(修正)', '活动名称', '活动名', '项目名称', '项目名', 'activity', 'activity_name'],
};

function resolveCol(rows, key) {
  if (!rows || !rows.length) return COLS[key];
  const header = Object.keys(rows[0] || {});
  if (header.includes(COLS[key])) return COLS[key];
  const aliases = COLS_ALIASES[key];
  if (!aliases) return COLS[key];
  for (const a of aliases) {
    const found = header.find((h) => h.trim() === a || h.replace(/\ufeff/g, '').trim() === a);
    if (found) return found;
  }
  if (key === 'date' && header.find((h) => /日期|周|date|week/i.test(h))) return header.find((h) => /日期|周|date|week/i.test(h));
  if (key === 'period' && header.find((h) => /统计周期|周期|period/i.test(h))) return header.find((h) => /统计周期|周期|period/i.test(h));
  if (key === 'l1' && header.find((h) => /一级/i.test(h))) return header.find((h) => /一级/i.test(h));
  if (key === 'l2' && header.find((h) => /二级/i.test(h))) return header.find((h) => /二级/i.test(h));
  if (key === 'uv' && header.find((h) => /祈愿|访问|uv/i.test(h))) return header.find((h) => /祈愿|访问|uv/i.test(h));
  if (key === 'drawUsers' && header.find((h) => /抽卡.*用户/i.test(h))) return header.find((h) => /抽卡.*用户/i.test(h));
  if (key === 'payUsers' && header.find((h) => /付费.*用户|付费抽卡/i.test(h))) return header.find((h) => /付费.*用户|付费抽卡/i.test(h));
  if (key === 'payAmount' && header.find((h) => /付费.*金额|付费金额|金额/i.test(h))) return header.find((h) => /付费.*金额|付费金额|金额/i.test(h));
  if (key === 'activity' && header.find((h) => /活动名称|活动名|项目名称|项目名/i.test(h))) return header.find((h) => /活动名称|活动名|项目名称|项目名/i.test(h));
  return COLS[key];
}

const OPS_ALIASES = new Set(['广告_资源投放', '广告资源投放', 'v2_资源投放', 'v2资源投放']);

const WISH_SUBDIR_CANDIDATES = ['祈愿', '祈愿bar来源'];

function num(v) {
  if (v == null) return 0;
  const s = String(v).trim();
  if (!s) return 0;
  const n = Number(s.replace(/,/g, ''));
  return Number.isFinite(n) ? n : 0;
}

function normSource(s) {
  const v = (s == null ? '' : String(s)).trim();
  if (!v || v === '--' || v.toLowerCase() === 'nan') return '(空)';
  if (OPS_ALIASES.has(v)) return '运营资源投放';
  return v;
}

/**
 * 周维度「日期」列统一为 YYYY-MM-DD。
 * 混用 2026/03/29 与 2026-03-23 时，原生字符串排序会把 '/' 格式整体排错，导致「最新周」仍停在 3/23。
 */
function normalizeWishbarWeekKey(raw) {
  const s = String(raw ?? '').replace(/\ufeff/g, '').trim();
  if (!s) return '';
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${String(m[2]).padStart(2, '0')}-${String(m[3]).padStart(2, '0')}`;
  m = s.match(/^(\d{4})[./](\d{1,2})[./](\d{1,2})/);
  if (m) return `${m[1]}-${String(m[2]).padStart(2, '0')}-${String(m[3]).padStart(2, '0')}`;
  m = s.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日?/);
  if (m) return `${m[1]}-${String(m[2]).padStart(2, '0')}-${String(m[3]).padStart(2, '0')}`;
  m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  return s;
}

function weekKeyToTimeMs(key) {
  const n = normalizeWishbarWeekKey(key);
  const m = n.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const t = d.getTime();
  return Number.isFinite(t) ? t : null;
}

function compareWeekKeysChronological(a, b) {
  const ta = weekKeyToTimeMs(a);
  const tb = weekKeyToTimeMs(b);
  if (ta != null && tb != null && ta !== tb) return ta - tb;
  return String(a).localeCompare(String(b), 'zh-CN');
}

/** 一级来源缺省合并为「(空)」：图表/宣推表/周明细列默认不展示 */
function isEmptyPrimarySource(label) {
  return label === '(空)';
}

/** 每千次访问带来的量（抽卡用户 / 付费抽卡用户 / 金额） */
function perThousandByVisits(numerator, uv) {
  if (uv <= 0 || !Number.isFinite(numerator)) return null;
  return (numerator / uv) * 1000;
}

function fmtPerKUsers(n, uv) {
  const v = perThousandByVisits(n, uv);
  if (v == null) return '—';
  return `${v.toFixed(2)} 人`;
}

function fmtPerKAmount(amt, uv) {
  const v = perThousandByVisits(amt, uv);
  if (v == null) return '—';
  if (v >= 1e4) return `${(v / 1e4).toFixed(2)} 万元`;
  return `${fmtNum(v)} 元`;
}

function pickMaxRow(rows, scoreFn) {
  const scored = rows
    .map((r) => ({ r, s: scoreFn(r) }))
    .filter((x) => Number.isFinite(x.s));
  if (!scored.length) return null;
  scored.sort((a, b) => b.s - a.s);
  return scored[0].r;
}

/** 模块二 / 周报：按周聚合的一级来源行 */
function buildWishbarContributionRows(fp, weekStr) {
  if (!fp || !weekStr) return [];
  const uvMap = fp.l1UvByWeek.get(weekStr) || new Map();
  const drawMap = fp.l1DrawByWeek.get(weekStr) || new Map();
  const payUserMap = fp.l1PayUserByWeek.get(weekStr) || new Map();
  const payAmountMap = fp.l1PayAmountByWeek.get(weekStr) || new Map();
  const shareMap = fp.l1SharesByWeek.get(weekStr) || new Map();
  const drawShareMap = fp.l1DrawShareByWeek?.get(weekStr) || new Map();
  const payUserShareMap = fp.l1PayUserShareByWeek?.get(weekStr) || new Map();
  const payAmountShareMap = fp.l1PayAmountShareByWeek?.get(weekStr) || new Map();

  return fp.l1All
    .filter((l1) => !isEmptyPrimarySource(l1))
    .map((l1) => ({
      l1,
      uv: uvMap.get(l1) || 0,
      draw: drawMap.get(l1) || 0,
      payUser: payUserMap.get(l1) || 0,
      payAmount: payAmountMap.get(l1) || 0,
      uvShare: shareMap.get(l1) || 0,
      drawShare: drawShareMap.get(l1) || 0,
      payUserShare: payUserShareMap.get(l1) || 0,
      payAmountShare: payAmountShareMap.get(l1) || 0,
    }))
    .filter((r) => r.uv > 0 || r.draw > 0 || r.payUser > 0 || r.payAmount > 0)
    .sort((a, b) => (b.uvShare || 0) - (a.uvShare || 0));
}

function parseWeekToDate(w) {
  const n = normalizeWishbarWeekKey(w);
  const m = n.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return null;
}

function calendarQuarterFromDate(d) {
  return Math.floor(d.getMonth() / 3) + 1;
}

/** 与 anchor 周同属自然年的季度（周标签需可解析为 YYYY-MM-DD） */
function filterWeeksSameQuarter(sortedWeeks, anchorWeek) {
  const d0 = parseWeekToDate(anchorWeek);
  if (!d0) return sortedWeeks.slice();
  const y = d0.getFullYear();
  const q = calendarQuarterFromDate(d0);
  return sortedWeeks.filter((w) => {
    const d = parseWeekToDate(w);
    if (!d) return false;
    return d.getFullYear() === y && calendarQuarterFromDate(d) === q;
  });
}

function linearRegressionSlopeY(values) {
  const n = values.length;
  if (n < 2) return null;
  let sumX = 0;
  let sumY = 0;
  let sumXY = 0;
  let sumXX = 0;
  for (let i = 0; i < n; i += 1) {
    sumX += i;
    sumY += values[i];
    sumXY += i * values[i];
    sumXX += i * i;
  }
  const denom = n * sumXX - sumX * sumX;
  if (denom === 0) return 0;
  return (n * sumXY - sumX * sumY) / denom;
}

function fmtWowRatio(cur, prev) {
  if (prev == null || !Number.isFinite(prev) || prev === 0) return null;
  if (!Number.isFinite(cur)) return null;
  return ((cur - prev) / prev) * 100;
}

function fmtWowArrow(pct) {
  if (pct == null || !Number.isFinite(pct)) return '暂无有效环比';
  if (pct > 0.5) return `环比上升约 ${pct.toFixed(1)}%`;
  if (pct < -0.5) return `环比下降约 ${Math.abs(pct).toFixed(1)}%`;
  return `环比基本持平（${pct >= 0 ? '+' : ''}${pct.toFixed(1)}%）`;
}

/** 简报 KPI 主数字后的周环比 HTML，如 <span>（+12%）</span>（上周为 0 或无上周时为 —） */
function fmtWowParenHtml(pct) {
  if (pct == null || !Number.isFinite(pct)) {
    return '<span class="wishbarReport__chipWow wishbarReport__chipWow--na">（—）</span>';
  }
  if (Math.abs(pct) < 0.05) {
    return '<span class="wishbarReport__chipWow wishbarReport__chipWow--flat">（0%）</span>';
  }
  const decimals = Math.abs(pct) >= 10 ? 0 : 1;
  const sign = pct > 0 ? '+' : '';
  const cls =
    pct > 0 ? 'wishbarReport__chipWow wishbarReport__chipWow--up' : 'wishbarReport__chipWow wishbarReport__chipWow--down';
  return `<span class="${cls}">（${sign}${pct.toFixed(decimals)}%）</span>`;
}

/**
 * 全站某指标周环比的一级来源归因：在「变动方向与全站一致」的来源中取绝对增量最大者；
 * 若无同向来源则回退为绝对增量最大者（排除一级「(空)」）。
 * @param {'uv'|'draw'|'payUser'|'payAmount'} field
 * @returns {{l1:string,delta:number,sourceWow:number|null,shareOfTotal:number|null}|null}
 */
function topAttributionL1(fp, weekStr, prevWeekStr, field) {
  const totalKeys = { uv: 'uv', draw: 'draw', payUser: 'payUser', payAmount: 'payAmount' };
  const mapSeries = {
    uv: fp.l1UvByWeek,
    draw: fp.l1DrawByWeek,
    payUser: fp.l1PayUserByWeek,
    payAmount: fp.l1PayAmountByWeek,
  }[field];
  const tk = totalKeys[field];
  if (!fp || !weekStr || !prevWeekStr || !mapSeries) return null;

  const curM = mapSeries.get(weekStr) || new Map();
  const prevM = mapSeries.get(prevWeekStr) || new Map();
  const curT = Number((fp.totalsByWeek.get(weekStr) || {})[tk] ?? 0);
  const prevT = Number((fp.totalsByWeek.get(prevWeekStr) || {})[tk] ?? 0);
  const totalDelta = curT - prevT;

  if (!Number.isFinite(totalDelta) || Math.abs(totalDelta) < 1e-9) return null;

  const consider = (sameSignOnly) => {
    let best = null;
    let bestAbs = -1;
    for (const l1 of fp.l1All || []) {
      if (isEmptyPrimarySource(l1)) continue;
      const cur = Number(curM.get(l1) ?? 0);
      const prev = Number(prevM.get(l1) ?? 0);
      const d = cur - prev;
      if (sameSignOnly && Math.sign(d) !== Math.sign(totalDelta)) continue;
      const ad = Math.abs(d);
      if (ad > bestAbs + 1e-9) {
        bestAbs = ad;
        best = { l1, d, cur, prev };
      } else if (best && Math.abs(ad - bestAbs) <= 1e-9 && l1.localeCompare(best.l1, 'zh-CN') < 0) {
        best = { l1, d, cur, prev };
      }
    }
    return best;
  };

  let pick = consider(true);
  if (!pick || Math.abs(pick.d) < 1e-9) pick = consider(false);
  if (!pick || Math.abs(pick.d) < 1e-9) return null;

  const sourceWow = fmtWowRatio(pick.cur, pick.prev);
  let shareOfTotal = null;
  if (Math.sign(pick.d) === Math.sign(totalDelta) && Math.abs(totalDelta) > 1e-9) {
    const r = pick.d / totalDelta;
    if (r > 0 && r <= 1.0001) shareOfTotal = Math.min(1, r);
  }

  return { l1: pick.l1, delta: pick.d, sourceWow, shareOfTotal };
}

/** KPI 卡片底部：周环比有方向时展示一级来源归因（内联 HTML，已 escape 名称） */
function fmtKpiAttributionHtml(fp, weekStr, prevWeekStr, field, totalWowPct) {
  if (!prevWeekStr || totalWowPct == null || !Number.isFinite(totalWowPct) || Math.abs(totalWowPct) < 0.05) return '';
  const att = topAttributionL1(fp, weekStr, prevWeekStr, field);
  if (!att) return '';
  const name = escapeHtml(att.l1);
  let parts = '';
  if (att.shareOfTotal != null && att.shareOfTotal > 0) {
    parts += ` · 约占全站变动 <strong>${Math.round(att.shareOfTotal * 100)}%</strong>`;
  }
  if (att.sourceWow != null && Number.isFinite(att.sourceWow) && Math.abs(att.sourceWow) >= 0.05) {
    const dec = Math.abs(att.sourceWow) >= 10 ? 0 : 1;
    const sg = att.sourceWow > 0 ? '+' : '';
    parts += ` · 该来源 <strong>${sg}${att.sourceWow.toFixed(dec)}%</strong>`;
  }
  return `<div class="wishbarReport__chipAttr">主因一级来源：<strong>${name}</strong>${parts}</div>`;
}

function rowCell(r, colName) {
  if (!r || colName == null) return '';
  if (r[colName] != null && r[colName] !== '') return r[colName];
  const bom = `\ufeff${colName}`;
  if (r[bom] != null && r[bom] !== '') return r[bom];
  return r[colName];
}

function activityNameFromRow(r, colAct) {
  const v = rowCell(r, colAct);
  const s = (v == null ? '' : String(v)).trim();
  return s || '(未命名活动)';
}

/** 简报「付费项目 TopN」：汇总行等活动名不参与聚合与排序（与活动列 trim 后全等匹配）。 */
const ACTIVITY_TOPN_EXCLUDE_NAMES = new Set(['整体']);

function isActivityExcludedFromTopN(name) {
  const n = String(name ?? '').trim();
  return ACTIVITY_TOPN_EXCLUDE_NAMES.has(n);
}

function headerHasColumn(rows, logicalKey) {
  if (!rows?.length) return false;
  const resolved = resolveCol(rows, logicalKey);
  return Object.keys(rows[0] || {}).some(
    (k) => k === resolved || k.replace(/\ufeff/g, '').trim() === resolved,
  );
}

/**
 * 当周按活动名称聚合，付费贡献 TopN。
 * 活动名为「整体」的行视为全表汇总，不参与聚合与排序。
 * 综合分 = 当周内「付费金额」「付费抽卡用户数」分别除以当周同列最大值归一化后，按 **金额 70% + 用户数 30%** 加权。
 */
function computeActivityPayTopN(rows, weekStr, topN = 5) {
  if (!rows?.length || !weekStr) {
    return { items: [], ok: false, hint: '无原始行或周为空。' };
  }
  if (!headerHasColumn(rows, 'activity')) {
    return {
      items: [],
      ok: false,
      hint: 'CSV 中未识别到「活动名称【修正】」或兼容列名（活动名称/项目名等），无法汇总项目维度。',
    };
  }

  const colDate = resolveCol(rows, 'date');
  const colPeriod = resolveCol(rows, 'period');
  const colAct = resolveCol(rows, 'activity');
  const colPayU = resolveCol(rows, 'payUsers');
  const colPayA = resolveCol(rows, 'payAmount');
  const hasPeriodCol = rows.some((r) => r && (r[colPeriod] != null && r[colPeriod] !== ''));

  const map = new Map();
  for (const r of rows) {
    if (!r) continue;
    const wk = normalizeWishbarWeekKey(String(rowCell(r, colDate) ?? '').trim());
    if (wk !== weekStr) continue;
    if (hasPeriodCol) {
      const period = String(rowCell(r, colPeriod) ?? '').trim();
      if (!period || !(period === '周' || period.startsWith('周') || period.includes('周'))) continue;
    }
    const name = activityNameFromRow(r, colAct);
    if (isActivityExcludedFromTopN(name)) continue;
    const pu = num(rowCell(r, colPayU));
    const pa = num(rowCell(r, colPayA));
    if (pu === 0 && pa === 0) continue;
    const cur = map.get(name) || { payUser: 0, payAmount: 0 };
    cur.payUser += pu;
    cur.payAmount += pa;
    map.set(name, cur);
  }

  const arr = Array.from(map.entries()).map(([name, v]) => ({ name, payUser: v.payUser, payAmount: v.payAmount }));
  if (!arr.length) {
    return { items: [], ok: true, hint: '当周无付费相关行（付费抽卡用户与金额均为 0），或周期过滤后无数据。' };
  }

  const maxU = Math.max(...arr.map((x) => x.payUser), 0);
  const maxA = Math.max(...arr.map((x) => x.payAmount), 0);
  const denomU = maxU > 0 ? maxU : 1;
  const denomA = maxA > 0 ? maxA : 1;
  const scored = arr.map((x) => ({
    name: x.name,
    payUser: x.payUser,
    payAmount: x.payAmount,
    score: 0.3 * (x.payUser / denomU) + 0.7 * (x.payAmount / denomA),
  }));
  scored.sort((a, b) => b.score - a.score);
  return { items: scored.slice(0, topN), ok: true, hint: '' };
}

function buildWishbarWeeklyReportHtml(fp, weekForM2, rows, rawRowsFiltered) {
  const minUvK = 1000;
  const dense = rows.filter((r) => r.uv >= minUvK);
  const topUvShare = pickMaxRow(rows, (r) => r.uvShare);
  const topPayUserSh = pickMaxRow(rows, (r) => r.payUserShare);
  const topPayAmtSh = pickMaxRow(rows, (r) => r.payAmountShare);
  const scorePerK = (num, uv) => {
    const v = perThousandByVisits(num, uv);
    return v == null ? NaN : v;
  };
  const topDrawK = pickMaxRow(dense, (r) => scorePerK(r.draw, r.uv));
  const topPayUserK = pickMaxRow(dense, (r) => scorePerK(r.payUser, r.uv));
  const topPayAmtK = pickMaxRow(dense, (r) => scorePerK(r.payAmount, r.uv));

  const totals = fp.totalsByWeek.get(weekForM2) || {};
  const tUv = totals.uv || 0;
  const tDraw = totals.draw || 0;
  const tPayUser = totals.payUser || 0;
  const tPayAmt = totals.payAmount || 0;

  const idx = fp.weeks.indexOf(weekForM2);
  const prevWeek = idx > 0 ? fp.weeks[idx - 1] : null;
  const prevTotals = prevWeek ? fp.totalsByWeek.get(prevWeek) || {} : null;
  const pUv = prevTotals?.uv;
  const pDraw = prevTotals?.draw;
  const pPayUser = prevTotals?.payUser;
  const pPayAmt = prevTotals?.payAmount;

  const wowUv = fmtWowRatio(tUv, pUv);
  const wowDraw = fmtWowRatio(tDraw, pDraw);
  const wowPayU = fmtWowRatio(tPayUser, pPayUser);
  const wowPayAmt = fmtWowRatio(tPayAmt, pPayAmt);

  const prevRows = prevWeek ? buildWishbarContributionRows(fp, prevWeek) : [];
  let shareLine = '';
  if (topUvShare && prevRows.length) {
    const pr = prevRows.find((x) => x.l1 === topUvShare.l1);
    if (pr) {
      const dpp = topUvShare.uvShare - pr.uvShare;
      shareLine = `「${escapeHtml(topUvShare.l1)}」访问占比 <strong>${dpp >= 0 ? '+' : ''}${dpp.toFixed(1)}</strong>pp`;
    }
  }

  const qWeeks = filterWeeksSameQuarter(fp.weeks, weekForM2);
  const uvSeries = qWeeks.map((w) => (fp.totalsByWeek.get(w) || {}).uv || 0);
  const slope = uvSeries.length >= 2 ? linearRegressionSlopeY(uvSeries) : null;
  const meanUv = uvSeries.length ? uvSeries.reduce((a, b) => a + b, 0) / uvSeries.length : 0;
  let qTrendLine = '';
  if (parseWeekToDate(weekForM2) && qWeeks.length >= 3 && slope != null && meanUv > 0) {
    const rel = slope / meanUv;
    const up = slope > 0;
    const notable = Math.abs(rel) > 0.03;
    qTrendLine = `本季 <strong>${qWeeks.length}</strong> 周 · 周UV拟合<strong>${up ? '偏升' : '走平/偏弱'}</strong>${notable ? ' · 波动明显' : ''}`;
  } else if (qWeeks.length >= 1) {
    const parsedOk = !!parseWeekToDate(weekForM2);
    qTrendLine = parsedOk
      ? `本季 <strong>${qWeeks.length}</strong> 周，可继续累积后再看斜率`
      : '周标签难解析日期，季内节奏请结合模块三';
  } else {
    qTrendLine = '—';
  }

  const top3 = rows
    .slice(0, 3)
    .map((r) => `<strong>${escapeHtml(r.l1)}</strong> ${fmtPct(r.uvShare)}`)
    .join(' · ');

  /** 付费相关「谁第一」：与来源Top3（按访问）区分，避免混在一行里难读 */
  const highlightItems = [];
  if (topPayUserSh) {
    highlightItems.push(
      `<li><span class="wishbarReport__hlMetric">付费用户占比</span> 最高 <strong>${escapeHtml(topPayUserSh.l1)}</strong> <span class="wishbarReport__hlVal">${fmtPct(topPayUserSh.payUserShare)}</span></li>`,
    );
  }
  if (topPayAmtSh) {
    highlightItems.push(
      `<li><span class="wishbarReport__hlMetric">付费金额占比</span> 最高 <strong>${escapeHtml(topPayAmtSh.l1)}</strong> <span class="wishbarReport__hlVal">${fmtPct(topPayAmtSh.payAmountShare)}</span></li>`,
    );
  }
  if (dense.length && topPayAmtK) {
    highlightItems.push(
      `<li><span class="wishbarReport__hlMetric">千次付费额</span> 最高（UV≥${minUvK}） <strong>${escapeHtml(topPayAmtK.l1)}</strong> <span class="wishbarReport__hlVal">${fmtPerKAmount(topPayAmtK.payAmount, topPayAmtK.uv)}</span></li>`,
    );
  } else if (!dense.length) {
    highlightItems.push(
      `<li><span class="wishbarReport__hlMetric">千次付费额</span> 各来源 UV 均 &lt; ${minUvK}，未做千次排名（见模块二表）</li>`,
    );
  }
  const highlightsListHtml = highlightItems.length
    ? `<ul class="wishbarReport__highlightsList">${highlightItems.join('')}</ul>`
    : '<p class="wishbarReport__muted">暂无一级来源付费结构明细。</p>';

  let highlightsDiverge = '';
  const uvL1 = topUvShare?.l1;
  const puL1 = topPayUserSh?.l1;
  const paL1 = topPayAmtSh?.l1;
  const kL1 = dense.length && topPayAmtK ? topPayAmtK.l1 : null;
  if (uvL1 && paL1 && uvL1 !== paL1) {
    highlightsDiverge = `<p class="wishbarReport__highlightsNote">对照：访问体量第一是「<strong>${escapeHtml(uvL1)}</strong>」，付费金额占比第一是「<strong>${escapeHtml(paL1)}</strong>」——金额更集中时可结合客单/品类看是否匹配。</p>`;
  } else if (uvL1 && puL1 && uvL1 !== puL1) {
    highlightsDiverge = `<p class="wishbarReport__highlightsNote">对照：访问第一是「<strong>${escapeHtml(uvL1)}</strong>」，付费用户占比第一是「<strong>${escapeHtml(puL1)}</strong>」——可关注该来源付费转化是否被低估。</p>`;
  } else if (uvL1 && kL1 && uvL1 !== kL1 && paL1 === uvL1) {
    highlightsDiverge = `<p class="wishbarReport__highlightsNote">体量主力「<strong>${escapeHtml(uvL1)}</strong>」与千次付费额最高「<strong>${escapeHtml(kL1)}</strong>」不同——效率优势来源可在成本可控下小步加量验证。</p>`;
  }

  const tips = [];
  if (topUvShare && topPayAmtK && topUvShare.l1 === topPayAmtK.l1) {
    tips.push(`主阵地「${escapeHtml(topUvShare.l1)}」量效均优，适合维稳投入。`);
  } else if (topUvShare && topPayAmtK && topUvShare.l1 !== topPayAmtK.l1) {
    tips.push(`体量「${escapeHtml(topUvShare.l1)}」、千次付费「${escapeHtml(topPayAmtK.l1)}」— 效率侧可小步加量，体量侧抓转化。`);
  }
  if (topPayUserK && topUvShare && topPayUserK.l1 !== topUvShare.l1) {
    tips.push(`「${escapeHtml(topPayUserK.l1)}」千次付费用户突出，可做拉新+付费实验。`);
  }
  if (wowUv != null && wowUv < -5) tips.push('访问环比偏弱，复盘入口/活动/异常。');
  else if (wowUv != null && wowUv > 5) tips.push('访问环比走强，可沉淀动作并适度加码健康来源。');
  if (!tips.length) tips.push('对照模块二、三，关注占比与千次是否背离。');
  const tipsPick = tips.slice(0, 3);

  const actTop = computeActivityPayTopN(rawRowsFiltered || [], weekForM2, 5);
  let activityBlock = '';
  if (!actTop.ok) {
    activityBlock = `<div class="wishbarReport__subblock"><h4 class="wishbarReport__subh">付费项目 Top5</h4><p class="wishbarReport__muted">${escapeHtml(actTop.hint)}</p></div>`;
  } else if (actTop.items.length) {
    const tr = actTop.items
      .map(
        (x, i) => `<tr><td class="wishbarReport__tdNum">${i + 1}</td><td>${escapeHtml(x.name)}</td><td class="mono">${fmtNum(x.payUser)}</td><td class="mono">${fmtNum(x.payAmount)}</td></tr>`,
      )
      .join('');
    activityBlock = `<div class="wishbarReport__subblock">
      <h4 class="wishbarReport__subh">付费项目 Top5 <span class="wishbarReport__tag">活动名称【修正】</span></h4>
      <div class="wishbarReport__tableScroll"><table class="wishbarReport__table"><thead><tr><th>#</th><th>活动</th><th>付费用户</th><th>付费金额</th></tr></thead><tbody>${tr}</tbody></table></div>
      <p class="wishbarReport__note">排序：周内归一后 <strong>金额70% + 付费用户30%</strong></p>
    </div>`;
  } else {
    activityBlock = `<div class="wishbarReport__subblock"><h4 class="wishbarReport__subh">付费项目 Top5</h4><p class="wishbarReport__muted">${escapeHtml(actTop.hint)}</p></div>`;
  }

  const wowRows = prevWeek
    ? `<tr><td>访问</td><td>${fmtWowArrow(wowUv)}</td><td class="mono wishbarReport__tdSub">${fmtNum(pUv || 0)} → ${fmtNum(tUv)}</td></tr>
      <tr><td>抽卡用户</td><td>${fmtWowArrow(wowDraw)}</td><td class="mono wishbarReport__tdSub">${fmtNum(pDraw || 0)} → ${fmtNum(tDraw)}</td></tr>
      <tr><td>付费用户</td><td>${fmtWowArrow(wowPayU)}</td><td class="mono wishbarReport__tdSub">${fmtNum(pPayUser || 0)} → ${fmtNum(tPayUser)}</td></tr>
      <tr><td>付费金额</td><td>${fmtWowArrow(wowPayAmt)}</td><td class="mono wishbarReport__tdSub">${fmtNum(pPayAmt || 0)} → ${fmtNum(tPayAmt)}</td></tr>`
    : '';

  return `
<div class="wishbarReport">
  <div class="wishbarReport__head">
    <span class="wishbarReport__week">${escapeHtml(weekForM2)}</span>
    <span class="wishbarReport__metaInline">全局筛选后合并</span>
  </div>

  <section class="wishbarReport__block">
    <div class="wishbarReport__kpiRow">
      <div class="wishbarReport__chip"><span class="wishbarReport__chipVal">${fmtNum(tUv)}${fmtWowParenHtml(wowUv)}</span><span class="wishbarReport__chipLab">访问 · 周环比</span>${fmtKpiAttributionHtml(fp, weekForM2, prevWeek, 'uv', wowUv)}</div>
      <div class="wishbarReport__chip"><span class="wishbarReport__chipVal">${fmtNum(tDraw)}${fmtWowParenHtml(wowDraw)}</span><span class="wishbarReport__chipLab">抽卡用户 · 周环比</span>${fmtKpiAttributionHtml(fp, weekForM2, prevWeek, 'draw', wowDraw)}</div>
      <div class="wishbarReport__chip"><span class="wishbarReport__chipVal">${fmtNum(tPayUser)}${fmtWowParenHtml(wowPayU)}</span><span class="wishbarReport__chipLab">付费用户 · 周环比</span>${fmtKpiAttributionHtml(fp, weekForM2, prevWeek, 'payUser', wowPayU)}</div>
      <div class="wishbarReport__chip"><span class="wishbarReport__chipVal">${fmtNum(tPayAmt)}${fmtWowParenHtml(wowPayAmt)}</span><span class="wishbarReport__chipLab">付费金额 · 周环比</span>${fmtKpiAttributionHtml(fp, weekForM2, prevWeek, 'payAmount', wowPayAmt)}</div>
    </div>
    <p class="wishbarReport__tight"><span class="wishbarReport__lab">来源Top3</span> ${top3 || '—'} <span class="wishbarReport__labHint">（按访问占比）</span></p>
    <div class="wishbarReport__highlights">
      <p class="wishbarReport__highlightsLead"><span class="wishbarReport__lab">付费结构快照</span> 各指标「第一名」是谁（不含访问占比，避免与 Top3 重复）</p>
      ${highlightsListHtml}
      ${highlightsDiverge}
    </div>
    ${activityBlock}
  </section>

  <section class="wishbarReport__block wishbarReport__block--split">
    <div class="wishbarReport__split">
      <div class="wishbarReport__splitCol">
        <h3 class="wishbarReport__h wishbarReport__h--sm">环比上周</h3>
        ${
          prevWeek
            ? `<p class="wishbarReport__tight wishbarReport__muted">${escapeHtml(prevWeek)} → ${escapeHtml(weekForM2)}</p>
        <table class="wishbarReport__table wishbarReport__table--tight"><thead><tr><th>指标</th><th>环比</th><th>上周→本周</th></tr></thead><tbody>${wowRows}</tbody></table>
        ${shareLine ? `<p class="wishbarReport__tight">${shareLine}</p>` : ''}`
            : '<p class="wishbarReport__muted">无更早一周</p>'
        }
      </div>
      <div class="wishbarReport__splitCol">
        <h3 class="wishbarReport__h wishbarReport__h--sm">季内节奏</h3>
        <p class="wishbarReport__tight">${qTrendLine}</p>
        <p class="wishbarReport__note">历史拟合，非预测</p>
      </div>
    </div>
  </section>

  <section class="wishbarReport__block wishbarReport__block--action">
    <h3 class="wishbarReport__h wishbarReport__h--sm">建议</h3>
    <ul class="wishbarReport__ul wishbarReport__ul--tight wishbarReport__ul--action">
      ${tipsPick.map((t) => `<li>${t}</li>`).join('')}
    </ul>
  </section>
</div>`;
}

function parseCsvFile(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (res) => resolve(res.data || []),
      error: (err) => reject(err),
    });
  });
}

function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(DB_STORE)) db.createObjectStore(DB_STORE);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function dbGet(key) {
  const db = await openDb();
  try {
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, 'readonly');
      const store = tx.objectStore(DB_STORE);
      const req = store.get(key);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  } finally {
    db.close();
  }
}

async function dbSet(key, value) {
  const db = await openDb();
  try {
    await new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, 'readwrite');
      const store = tx.objectStore(DB_STORE);
      const req = store.put(value, key);
      req.onsuccess = () => resolve();
      req.onerror = () => reject(req.error);
    });
  } finally {
    db.close();
  }
}

async function getBoundDirHandle() {
  try {
    const h = await dbGet(DB_KEY_DIR_HANDLE);
    return h || null;
  } catch (_) {
    return null;
  }
}

async function setBoundDirHandle(handle) {
  await dbSet(DB_KEY_DIR_HANDLE, handle);
}

async function pickAndBindFolder() {
  if (!window.showDirectoryPicker) {
    throw new Error('当前浏览器不支持“绑定文件夹”（需要 File System Access API）。');
  }
  if (window.top && window.top !== window.self) {
    throw new Error('当前页面在汇总页(iframe)内，请在新窗口打开「祈愿bar访问贡献」页再绑定。');
  }
  const handle = await window.showDirectoryPicker({ mode: 'read' });
  await setBoundDirHandle(handle);
  return handle;
}

async function resolveWishSubdir(rootHandle) {
  for (const name of WISH_SUBDIR_CANDIDATES) {
    try {
      const h = await rootHandle.getDirectoryHandle(name, { create: false });
      return { handle: h, name };
    } catch (_) {}
  }
  const list = WISH_SUBDIR_CANDIDATES.map((x) => `「${x}」`).join(' 或 ');
  throw new Error(`未找到祈愿数据子文件夹（需在绑定目录下创建 ${list}）。`);
}

async function parseCsvText(text) {
  return new Promise((resolve, reject) => {
    Papa.parse(text, {
      header: true,
      skipEmptyLines: true,
      complete: (res) => resolve(res.data || []),
      error: (err) => reject(err),
    });
  });
}

function fmtPct(v) {
  if (!Number.isFinite(v)) return '0.00%';
  return `${v.toFixed(2)}%`;
}

function fmtNum(v) {
  if (!Number.isFinite(v)) return '0';
  if (v >= 10000) return (v / 10000).toFixed(2) + '万';
  return v.toLocaleString();
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/"/g, '&quot;');
}

/** option value 安全包裹（支持周字段含引号等） */
function optAttrValue(s) {
  return JSON.stringify(String(s));
}

let dropdownWeekSigFull = '';
let dropdownWeekSigFp = '';

let trendDateSig = '';

function resetWeekDropdownCache() {
  dropdownWeekSigFull = '';
  dropdownWeekSigFp = '';
  trendDateSig = '';
}

/** 仅截取指定周列表（用于模块三趋势图横轴） */
function sliceFpToWeekList(fp, weekList) {
  if (!fp || !weekList?.length) return fp;
  const pick = (map) => {
    const m = new Map();
    for (const w of weekList) {
      if (map?.has(w)) m.set(w, map.get(w));
    }
    return m;
  };
  return {
    ...fp,
    weeks: weekList.slice(),
    totalsByWeek: pick(fp.totalsByWeek),
    l1SharesByWeek: pick(fp.l1SharesByWeek),
    l2SharesByWeek: pick(fp.l2SharesByWeek),
    l1DrawShareByWeek: pick(fp.l1DrawShareByWeek || new Map()),
    l1PayUserShareByWeek: pick(fp.l1PayUserShareByWeek || new Map()),
    l1PayAmountShareByWeek: pick(fp.l1PayAmountShareByWeek || new Map()),
    l1UvByWeek: pick(fp.l1UvByWeek),
    l1DrawByWeek: pick(fp.l1DrawByWeek || new Map()),
    l1PayUserByWeek: pick(fp.l1PayUserByWeek || new Map()),
    l1PayAmountByWeek: pick(fp.l1PayAmountByWeek || new Map()),
  };
}

/** 模块三趋势：在全局筛选后的周序列上，再按起止周截取 */
function getTrendChartWeeks(fp) {
  const wf = fp?.weeks || [];
  const ss = $('wishbarTrendWeekStart');
  const se = $('wishbarTrendWeekEnd');
  if (!wf.length) return [];
  let ts = ss?.value || '';
  let te = se?.value || '';
  if (!ts) ts = wf[0];
  if (!te) te = wf[wf.length - 1];
  if (ts && te && ts > te) {
    const x = ts;
    ts = te;
    te = x;
  }
  const out = wf.filter((w) => w >= ts && w <= te);
  return out.length ? out : wf.slice();
}

function syncTrendChartWeekDropdowns(parsed, fp) {
  const full = parsed?.weeks || [];
  const ss = $('wishbarTrendWeekStart');
  const se = $('wishbarTrendWeekEnd');
  if (!ss || !se) return;

  const sig = full.join('\u0001');
  if (sig !== trendDateSig) {
    trendDateSig = sig;
    const head = '<option value="">不限</option>';
    const opts = full.map((w) => `<option value=${optAttrValue(w)}>${escapeHtml(w)}</option>`).join('');
    ss.innerHTML = head + opts;
    se.innerHTML = head + opts;
    const wf = fp?.weeks || [];
    if (wf.length) {
      ss.value = wf[0];
      se.value = wf[wf.length - 1];
    } else {
      ss.value = '';
      se.value = '';
    }
    return;
  }

  const wf = fp?.weeks || [];
  const inWf = (v) => !v || wf.includes(v);
  if (ss.value && !inWf(ss.value)) ss.value = wf[0] || '';
  if (se.value && !inWf(se.value)) se.value = wf[wf.length - 1] || '';
}

/** 占比序列 → 以首点为 100 的指数，便于看四条线的相对起伏 */
function toIndexFromFirst(pctArr) {
  if (!pctArr?.length) return [];
  const b = pctArr[0];
  if (!b || b === 0) return pctArr.map(() => 0);
  return pctArr.map((v) => (v / b) * 100);
}

/** 全局时间范围：选项来自完整解析结果中的「周」键，与 CSV 完全一致 */
function syncGlobalWeekDropdowns(parsed) {
  const weeksFull = parsed?.weeks || [];
  const startSel = $('wishbarDateStart');
  const endSel = $('wishbarDateEnd');
  if (!startSel || !endSel) return;

  const sigFull = weeksFull.join('\u0001');
  if (sigFull === dropdownWeekSigFull) return;

  dropdownWeekSigFull = sigFull;
  const head = '<option value="">不限</option>';
  const opts = weeksFull.map((w) => `<option value=${optAttrValue(w)}>${escapeHtml(w)}</option>`).join('');
  startSel.innerHTML = head + opts;
  endSel.innerHTML = head + opts;
  if (weeksFull.length) {
    startSel.value = weeksFull[0];
    endSel.value = weeksFull[weeksFull.length - 1];
  } else {
    startSel.value = '';
    endSel.value = '';
  }
}

/** 模块一/二：仅展示当前筛选后仍存在的周 */
function syncModuleWeekDropdowns(fp) {
  const weeksF = fp?.weeks || [];
  const w1 = $('wishbarWeekPick');
  const w2 = $('wishbarModule2Week');
  if (!w1 || !w2) return;

  const sigFp = weeksF.join('\u0001');
  if (sigFp !== dropdownWeekSigFp) {
    dropdownWeekSigFp = sigFp;
    const prev1 = w1.value;
    const prev2 = w2.value;
    const opts2 = weeksF.map((w) => `<option value=${optAttrValue(w)}>${escapeHtml(w)}</option>`).join('');
    w1.innerHTML = opts2;
    w2.innerHTML = opts2;
    const last = weeksF.length ? weeksF[weeksF.length - 1] : '';
    w1.value = weeksF.includes(prev1) ? prev1 : last;
    w2.value = weeksF.includes(prev2) ? prev2 : last;
  } else if (weeksF.length) {
    if (!weeksF.includes(w1.value)) w1.value = weeksF[weeksF.length - 1];
    if (!weeksF.includes(w2.value)) w2.value = weeksF[weeksF.length - 1];
  } else {
    w1.innerHTML = '';
    w2.innerHTML = '';
  }
}

function buildWeeklyShares(rows) {
  const colDate = resolveCol(rows, 'date');
  const colPeriod = resolveCol(rows, 'period');
  const colL1 = resolveCol(rows, 'l1');
  const colL2 = resolveCol(rows, 'l2');
  const colUv = resolveCol(rows, 'uv');
  const colDraw = resolveCol(rows, 'drawUsers');
  const colPayUser = resolveCol(rows, 'payUsers');
  const colPayAmount = resolveCol(rows, 'payAmount');

  const hasPeriodCol = rows.some((r) => r && (r[colPeriod] != null && r[colPeriod] !== ''));

  const weekData = new Map(); // week -> Map(l1 -> {uv, draw, payUser, payAmount})
  const weekL2Uv = new Map();

  for (const r of rows) {
    if (!r) continue;
    const periodRaw = r[colPeriod];
    const period = String(periodRaw ?? '').trim();
    if (hasPeriodCol && (!period || !(period === '周' || period.startsWith('周') || period.includes('周')))) continue;

    const week = normalizeWishbarWeekKey(String(r[colDate] ?? '').trim());
    if (!week) continue;

    const l1 = normSource(r[colL1]);
    const l2 = normSource(r[colL2]);
    const uv = num(r[colUv]);
    const draw = num(r[colDraw]);
    const payUser = num(r[colPayUser]);
    const payAmount = num(r[colPayAmount]);

    if (uv === 0 && draw === 0 && payUser === 0 && payAmount === 0) continue;

    if (!weekData.has(week)) {
      weekData.set(week, new Map());
    }
    const m = weekData.get(week);
    if (!m.has(l1)) m.set(l1, { uv: 0, draw: 0, payUser: 0, payAmount: 0 });
    const cur = m.get(l1);
    cur.uv += uv;
    cur.draw += draw;
    cur.payUser += payUser;
    cur.payAmount += payAmount;

    if (!weekL2Uv.has(week)) weekL2Uv.set(week, new Map());
    const m2 = weekL2Uv.get(week);
    m2.set(l2, (m2.get(l2) || 0) + uv);
  }

  const weeks = Array.from(weekData.keys()).sort(compareWeekKeysChronological);
  const l1Set = new Set();
  const l2Set = new Set();
  for (const w of weeks) {
    for (const s of weekData.get(w).keys()) l1Set.add(s);
    for (const s of (weekL2Uv.get(w) || new Map()).keys()) l2Set.add(s);
  }
  const l1All = Array.from(l1Set);
  const l2All = Array.from(l2Set);

  const l1UvByWeek = new Map();
  const l1DrawByWeek = new Map();
  const l1PayUserByWeek = new Map();
  const l1PayAmountByWeek = new Map();
  const totalsByWeek = new Map();
  const l1SharesByWeek = new Map();
  const l1DrawShareByWeek = new Map();
  const l1PayUserShareByWeek = new Map();
  const l1PayAmountShareByWeek = new Map();
  const l2SharesByWeek = new Map();

  for (const w of weeks) {
    const m = weekData.get(w);
    let tUv = 0, tDraw = 0, tPayUser = 0, tPayAmount = 0;
    const uvMap = new Map();
    const drawMap = new Map();
    const payUserMap = new Map();
    const payAmountMap = new Map();
    for (const [l1, d] of m.entries()) {
      uvMap.set(l1, d.uv);
      drawMap.set(l1, d.draw);
      payUserMap.set(l1, d.payUser);
      payAmountMap.set(l1, d.payAmount);
      tUv += d.uv;
      tDraw += d.draw;
      tPayUser += d.payUser;
      tPayAmount += d.payAmount;
    }
    l1UvByWeek.set(w, uvMap);
    l1DrawByWeek.set(w, drawMap);
    l1PayUserByWeek.set(w, payUserMap);
    l1PayAmountByWeek.set(w, payAmountMap);
    totalsByWeek.set(w, { uv: tUv, draw: tDraw, payUser: tPayUser, payAmount: tPayAmount });

    const s1 = new Map();
    const sDraw = new Map();
    const sPayUser = new Map();
    const sPayAmount = new Map();
    for (const s of l1All) {
      s1.set(s, tUv > 0 ? (uvMap.get(s) || 0) / tUv * 100 : 0);
      sDraw.set(s, tDraw > 0 ? (drawMap.get(s) || 0) / tDraw * 100 : 0);
      sPayUser.set(s, tPayUser > 0 ? (payUserMap.get(s) || 0) / tPayUser * 100 : 0);
      sPayAmount.set(s, tPayAmount > 0 ? (payAmountMap.get(s) || 0) / tPayAmount * 100 : 0);
    }
    l1SharesByWeek.set(w, s1);
    l1DrawShareByWeek.set(w, sDraw);
    l1PayUserShareByWeek.set(w, sPayUser);
    l1PayAmountShareByWeek.set(w, sPayAmount);

    const m2 = weekL2Uv.get(w) || new Map();
    let t2 = 0;
    for (const v of m2.values()) t2 += v;
    const s2 = new Map();
    for (const s of l2All) s2.set(s, t2 > 0 ? (m2.get(s) || 0) / t2 * 100 : 0);
    l2SharesByWeek.set(w, s2);
  }

  return {
    weeks,
    l1All,
    l2All,
    l1SharesByWeek,
    l2SharesByWeek,
    l1DrawShareByWeek,
    l1PayUserShareByWeek,
    l1PayAmountShareByWeek,
    l1UvByWeek,
    l1DrawByWeek,
    l1PayUserByWeek,
    l1PayAmountByWeek,
    l2UvByWeek: weekL2Uv,
    totalsByWeek,
  };
}

function getFilteredRows(rows, filters) {
  if (!rows || !rows.length) return rows;
  const colDate = resolveCol(rows, 'date');
  const colL1 = resolveCol(rows, 'l1');
  const colL2 = resolveCol(rows, 'l2');

  return rows.filter((r) => {
    const week = normalizeWishbarWeekKey(String(r[colDate] ?? '').trim());
    if (!week) return false;
    if (filters.dateStart && week < filters.dateStart) return false;
    if (filters.dateEnd && week > filters.dateEnd) return false;
    const l1 = normSource(r[colL1]);
    const l2 = normSource(r[colL2]);
    if (filters.l1Set.size && !filters.l1Set.has(l1)) return false;
    if (filters.l2Set.size && !filters.l2Set.has(l2)) return false;
    return true;
  });
}

/** @deprecated 当前未调用。全局筛选应使用 getFilteredParsed()：对 rawRows 做 getFilteredRows 后再 buildWeeklyShares，避免仅切片 map 导致占比分母与分子不一致。 */
function applyGlobalFilters(parsed, filters) {
  if (!parsed || !parsed.weeks.length) return parsed;
  let weeks = parsed.weeks;
  if (filters.dateStart) weeks = weeks.filter((w) => w >= filters.dateStart);
  if (filters.dateEnd) weeks = weeks.filter((w) => w <= filters.dateEnd);
  if (weeks.length === 0) return { ...parsed, weeks: [], l1All: [], l2All: [] };

  const l1Ok = (s) => !filters.l1Set.size || filters.l1Set.has(s);
  const l2Ok = (s) => !filters.l2Set.size || filters.l2Set.has(s);

  const l1All = parsed.l1All.filter(l1Ok);
  const l2All = parsed.l2All.filter(l2Ok);

  const slice = (map) => {
    const m = new Map();
    for (const w of weeks) if (map.has(w)) m.set(w, map.get(w));
    return m;
  };

  const sliceL1Map = (map) => {
    const out = new Map();
    for (const w of weeks) {
      const inner = map.get(w);
      if (!inner) continue;
      const filtered = new Map();
      for (const [k, v] of inner) if (l1Ok(k)) filtered.set(k, v);
      out.set(w, filtered);
    }
    return out;
  };

  const sliceL2Map = (map) => {
    const out = new Map();
    for (const w of weeks) {
      const inner = map.get(w);
      if (!inner) continue;
      const filtered = new Map();
      for (const [k, v] of inner) if (l2Ok(k)) filtered.set(k, v);
      out.set(w, filtered);
    }
    return out;
  };

  return {
    ...parsed,
    weeks,
    l1All,
    l2All,
    l1SharesByWeek: sliceL1Map(parsed.l1SharesByWeek),
    l2SharesByWeek: sliceL2Map(parsed.l2SharesByWeek),
    l1DrawShareByWeek: sliceL1Map(parsed.l1DrawShareByWeek || new Map()),
    l1PayUserShareByWeek: sliceL1Map(parsed.l1PayUserShareByWeek || new Map()),
    l1PayAmountShareByWeek: sliceL1Map(parsed.l1PayAmountShareByWeek || new Map()),
    l1UvByWeek: sliceL1Map(parsed.l1UvByWeek),
    l1DrawByWeek: sliceL1Map(parsed.l1DrawByWeek || new Map()),
    l1PayUserByWeek: sliceL1Map(parsed.l1PayUserByWeek || new Map()),
    l1PayAmountByWeek: sliceL1Map(parsed.l1PayAmountByWeek || new Map()),
    totalsByWeek: slice(parsed.totalsByWeek),
  };
}

function palette(n) {
  const base = ['#4f7cff', '#40c79a', '#f59e0b', '#ef4444', '#a855f7', '#06b6d4', '#22c55e', '#e11d48', '#0ea5e9', '#84cc16', '#f97316', '#64748b'];
  const out = [];
  for (let i = 0; i < n; i++) out.push(base[i % base.length]);
  return out;
}

function summarizeSelection(set, maxItems = 3) {
  const arr = Array.from(set);
  if (arr.length === 0) return '全部';
  const head = arr.slice(0, maxItems).join('、');
  return arr.length > maxItems ? `${head} 等 ${arr.length} 项` : `${head}（${arr.length}）`;
}

function ensureFilterPopover() {
  let pop = document.getElementById('filterPopover');
  if (pop) return pop;
  pop = document.createElement('div');
  pop.id = 'filterPopover';
  pop.className = 'popover';
  pop.innerHTML = `
    <div class="popover__head">
      <div class="popover__title" id="filterPopTitle">筛选</div>
      <button class="btn btn--ghost btn--sm" id="filterPopClose" type="button">关闭</button>
    </div>
    <div class="popover__body">
      <div class="filterSearch">
        <span style="color:rgba(15,23,42,.55)">🔎</span>
        <input id="filterPopSearch" type="text" placeholder="搜索…" />
      </div>
      <div class="filterRow" style="margin-top:10px">
        <button class="btn btn--ghost btn--sm" id="filterPopSelectAll" type="button">全选</button>
        <button class="btn btn--ghost btn--sm" id="filterPopInvert" type="button">反选</button>
        <button class="btn btn--ghost btn--sm" id="filterPopClear" type="button">清空</button>
        <span class="filterRow__right" id="filterPopCount">已选：0</span>
      </div>
      <div class="filterList" id="filterPopList"></div>
    </div>
    <div class="popover__footer">
      <button class="btn btn--ghost" id="filterPopCancel" type="button">取消</button>
      <button class="btn" id="filterPopOk" type="button">确定</button>
    </div>
  `;
  document.body.appendChild(pop);
  return pop;
}

function openFilterPopover({ anchorEl, title, items, selectedSet, onApply }) {
  const pop = ensureFilterPopover();
  const titleEl = document.getElementById('filterPopTitle');
  const closeBtn = document.getElementById('filterPopClose');
  const cancelBtn = document.getElementById('filterPopCancel');
  const okBtn = document.getElementById('filterPopOk');
  const searchEl = document.getElementById('filterPopSearch');
  const listEl = document.getElementById('filterPopList');
  const countEl = document.getElementById('filterPopCount');
  const btnAll = document.getElementById('filterPopSelectAll');
  const btnInv = document.getElementById('filterPopInvert');
  const btnClr = document.getElementById('filterPopClear');

  const temp = new Set(Array.from(selectedSet));
  let q = '';

  function visibleItems() {
    const qq = q.trim().toLowerCase();
    return qq ? items.filter((x) => String(x).toLowerCase().includes(qq)) : items;
  }
  function updateCount() {
    countEl.textContent = `已选：${temp.size}`;
  }
  function renderList() {
    const vis = visibleItems();
    listEl.innerHTML = vis
      .map((s) => {
        const checked = temp.has(s) ? 'checked' : '';
        const safe = String(s).replace(/</g, '&lt;').replace(/>/g, '&gt;');
        return `<label class="filterItem"><input type="checkbox" value="${safe}" ${checked} />${safe}</label>`;
      })
      .join('');
    listEl.querySelectorAll('input[type="checkbox"]').forEach((el) => {
      el.addEventListener('change', () => {
        if (el.checked) temp.add(el.value);
        else temp.delete(el.value);
        updateCount();
      });
    });
  }

  function close() {
    pop.style.display = 'none';
    document.removeEventListener('mousedown', onOutsideDown, true);
    document.removeEventListener('keydown', onKey, true);
  }
  function onKey(e) { if (e.key === 'Escape') close(); }
  function onOutsideDown(e) {
    if (!pop.contains(e.target) && e.target !== anchorEl) close();
  }

  titleEl.textContent = title;
  searchEl.value = '';
  q = '';
  updateCount();
  renderList();

  closeBtn.onclick = close;
  cancelBtn.onclick = close;
  okBtn.onclick = () => { onApply(new Set(temp)); close(); };
  btnClr.onclick = () => { temp.clear(); updateCount(); renderList(); };
  btnAll.onclick = () => { for (const s of visibleItems()) temp.add(s); updateCount(); renderList(); };
  btnInv.onclick = () => {
    const vis = visibleItems();
    for (const s of vis) { if (temp.has(s)) temp.delete(s); else temp.add(s); }
    updateCount();
    renderList();
  };
  searchEl.oninput = () => { q = searchEl.value || ''; renderList(); };

  const rect = anchorEl.getBoundingClientRect();
  const vw = Math.max(document.documentElement.clientWidth || 0, window.innerWidth || 0);
  const vh = Math.max(document.documentElement.clientHeight || 0, window.innerHeight || 0);
  const width = Math.min(720, vw - 24);
  let left = Math.max(12, rect.left);
  left = Math.min(left, vw - width - 12);
  let top = rect.bottom + 8;
  if (top + 420 > vh - 12) top = Math.max(12, rect.top - 8 - 420);

  pop.style.width = `${width}px`;
  pop.style.left = `${left}px`;
  pop.style.top = `${top}px`;
  pop.style.display = 'block';

  setTimeout(() => searchEl.focus(), 0);
  document.addEventListener('mousedown', onOutsideDown, true);
  document.addEventListener('keydown', onKey, true);
}

let chart = null;
let volumeChart = null;
let parsed = null;
let rawRows = [];
/** 汇总页快照导出：let 不会挂 window，供父页 iframe 读取 */
if (typeof window !== 'undefined') {
  Object.defineProperty(window, '__SNAPSHOT_WISHBAR_ROWS__', {
    configurable: true,
    enumerable: true,
    get() {
      return rawRows;
    },
  });
}
let globalL1Set = new Set();
let globalL2Set = new Set();
let trendSelectedL1 = new Set();
let trendInitialized = false;

function setStatus(msg) {
  const el = $('wishbarStatusHint');
  if (el) el.textContent = msg;
}

function getGlobalFilters() {
  const start = $('wishbarDateStart')?.value || '';
  const end = $('wishbarDateEnd')?.value || '';
  return {
    dateStart: start || null,
    dateEnd: end || null,
    l1Set: globalL1Set,
    l2Set: globalL2Set,
  };
}

function getFilteredParsed() {
  if (!parsed) return null;
  const filters = getGlobalFilters();
  const hasRowFilter = filters.dateStart || filters.dateEnd || filters.l1Set.size || filters.l2Set.size;
  if (hasRowFilter) {
    const fr = getFilteredRows(rawRows, filters);
    return buildWeeklyShares(fr);
  }
  return parsed;
}

/** 与全局筛选一致、用于模块三并集汇总的原始行 */
function getRowsForTrendAggregate() {
  if (!rawRows?.length) return [];
  const filters = getGlobalFilters();
  const hasRowFilter = filters.dateStart || filters.dateEnd || filters.l1Set.size || filters.l2Set.size;
  return hasRowFilter ? getFilteredRows(rawRows, filters) : rawRows;
}

/**
 * 按周汇总：一级并集（+ 可选二级）命中行的指标 ÷ 当周全量（fp.totalsByWeek，与 buildWeeklyShares 一致）。
 * 分母 t.* 为当周所有进入周聚合的行（含「(空)」一级），非仅选中来源。
 * 模块三仅传 trendL2 为空 Set。
 */
function aggregateTrendUnionSeries(rows, fp, trendL1, trendL2) {
  if (!rows?.length || !fp?.weeks?.length) {
    return {
      uvPct: [], drawPct: [], payUserPct: [], payAmtPct: [],
      uvRaw: [], drawRaw: [], payUserRaw: [], payAmtRaw: [],
    };
  }
  const colDate = resolveCol(rows, 'date');
  const colPeriod = resolveCol(rows, 'period');
  const colL1 = resolveCol(rows, 'l1');
  const colL2 = resolveCol(rows, 'l2');
  const colUv = resolveCol(rows, 'uv');
  const colDraw = resolveCol(rows, 'drawUsers');
  const colPayUser = resolveCol(rows, 'payUsers');
  const colPayAmount = resolveCol(rows, 'payAmount');
  const hasPeriodCol = rows.some((r) => r && (r[colPeriod] != null && r[colPeriod] !== ''));

  const cleanL1 = Array.from(trendL1).filter((x) => x !== '其他');
  const hasOtherL1 = trendL1.has('其他');

  function rowInUnion(l1, l2) {
    if (!trendL1.size && !trendL2.size) return true;
    let hitL1 = false;
    if (trendL1.size) {
      if (trendL1.has(l1)) hitL1 = true;
      else if (hasOtherL1 && !cleanL1.includes(l1)) hitL1 = true;
    }
    const hitL2 = trendL2.size > 0 && trendL2.has(l2);
    if (trendL1.size && !trendL2.size) return hitL1;
    if (!trendL1.size && trendL2.size) return hitL2;
    if (!trendL1.size && !trendL2.size) return true;
    return hitL1 || hitL2;
  }

  const uvPct = [];
  const drawPct = [];
  const payUserPct = [];
  const payAmtPct = [];
  const uvRaw = [];
  const drawRaw = [];
  const payUserRaw = [];
  const payAmtRaw = [];

  for (const w of fp.weeks) {
    let su = 0;
    let sd = 0;
    let spu = 0;
    let spa = 0;
    for (const r of rows) {
      if (String(r[colDate] ?? '').trim() !== w) continue;
      if (hasPeriodCol) {
        const period = String(r[colPeriod] ?? '').trim();
        if (!period || !(period === '周' || period.startsWith('周') || period.includes('周'))) continue;
      }
      const l1 = normSource(r[colL1]);
      const l2 = normSource(r[colL2]);
      if (!rowInUnion(l1, l2)) continue;
      su += num(r[colUv]);
      sd += num(r[colDraw]);
      spu += num(r[colPayUser]);
      spa += num(r[colPayAmount]);
    }
    const t = fp.totalsByWeek.get(w) || { uv: 0, draw: 0, payUser: 0, payAmount: 0 };
    uvRaw.push(su);
    drawRaw.push(sd);
    payUserRaw.push(spu);
    payAmtRaw.push(spa);
    uvPct.push(t.uv > 0 ? (su / t.uv) * 100 : 0);
    drawPct.push(t.draw > 0 ? (sd / t.draw) * 100 : 0);
    payUserPct.push(t.payUser > 0 ? (spu / t.payUser) * 100 : 0);
    payAmtPct.push(t.payAmount > 0 ? (spa / t.payAmount) * 100 : 0);
  }

  return {
    uvPct, drawPct, payUserPct, payAmtPct,
    uvRaw, drawRaw, payUserRaw, payAmtRaw,
  };
}

function buildSourceOrderByLatest(p, shareKey = 'l1SharesByWeek') {
  if (!p || !p.weeks.length) return [];
  const latest = p.weeks[p.weeks.length - 1];
  const m = p[shareKey]?.get(latest) || new Map();
  return p.l1All
    .filter((s) => !isEmptyPrimarySource(s))
    .map((s) => ({ s, v: m.get(s) || 0 }))
    .sort((a, b) => b.v - a.v)
    .map((x) => x.s);
}

function buildL2OrderByLatest(p) {
  if (!p || !p.weeks.length) return [];
  const latest = p.weeks[p.weeks.length - 1];
  const m = p.l2SharesByWeek?.get(latest) || new Map();
  return p.l2All
    .map((s) => ({ s, v: m.get(s) || 0 }))
    .sort((a, b) => b.v - a.v)
    .map((x) => x.s);
}

function initTrendDefaults(p) {
  if (trendInitialized) return;
  // 默认：全部非空一级来源并集 = 全站周汇总口径（占比曲线为 100%，与「看大盘」一致；可在一级来源里改选子集看结构）
  trendSelectedL1 = new Set(p.l1All.filter((s) => !isEmptyPrimarySource(s)));
  trendInitialized = true;
}

function render() {
  if (!parsed || parsed.weeks.length === 0) {
    const hint = rawRows?.length
      ? `已载入 ${rawRows.length} 行，但无符合条件数据。请确认 CSV 包含：日期、统计周期含「周」、一级来源、祈愿bar访问uv。当前表头：${rawRows[0] ? Object.keys(rawRows[0]).join('、') : ''}`
      : '请先选择 CSV 文件或绑定数据文件夹，再点击「读取祈愿文件夹全部CSV并更新」。';
    setStatus(hint);
    resetWeekDropdownCache();
    const ds = $('wishbarDateStart');
    const de = $('wishbarDateEnd');
    if (ds) ds.innerHTML = '';
    if (de) de.innerHTML = '';
    syncModuleWeekDropdowns({ weeks: [] });
    $('wishbarShareTable').innerHTML = '';
    const ctEmpty = $('wishbarContributionTable')?.querySelector('tbody');
    if (ctEmpty) ctEmpty.innerHTML = '';
    const ins = $('wishbarContributionInsight');
    if (ins) ins.textContent = '';
    if (chart) { chart.destroy(); chart = null; }
    if (volumeChart) { volumeChart.destroy(); volumeChart = null; }
    return;
  }

  syncGlobalWeekDropdowns(parsed);
  const fp = getFilteredParsed();
  if (!fp || fp.weeks.length === 0) {
    syncModuleWeekDropdowns({ weeks: [] });
    setStatus('当前筛选条件下无数据，请调整周范围或来源筛选。');
    $('wishbarShareTable').innerHTML = '';
    const ctEmpty2 = $('wishbarContributionTable')?.querySelector('tbody');
    if (ctEmpty2) ctEmpty2.innerHTML = '';
    const ins2 = $('wishbarContributionInsight');
    if (ins2) ins2.textContent = '';
    if (chart) { chart.destroy(); chart = null; }
    if (volumeChart) { volumeChart.destroy(); volumeChart = null; }
    return;
  }

  syncModuleWeekDropdowns(fp);
  syncTrendChartWeekDropdowns(parsed, fp);

  initTrendDefaults(fp);
  const l1TrendSum = $('wishbarTrendL1Summary');
  if (l1TrendSum) {
    l1TrendSum.textContent = trendSelectedL1.size
      ? summarizeSelection(trendSelectedL1)
      : '请先在一级来源中选至少一项';
  }

  const latestWeek = fp.weeks[fp.weeks.length - 1];
  const topX = parseInt($('wishbarModule1TopX')?.value || '5', 10);
  const weekPick = $('wishbarWeekPick')?.value || latestWeek;
  const weekForBar = fp.weeks.includes(weekPick) ? weekPick : latestWeek;

  // ----- 模块1：柱状图 TopX -----
  const l1Map = fp.l1SharesByWeek.get(weekForBar) || new Map();
  const barRaw = Array.from(l1Map.entries())
    .map(([label, value]) => ({ label, value }))
    .filter((x) => !isEmptyPrimarySource(x.label))
    .sort((a, b) => b.value - a.value);
  const barSeries = topX >= 999 ? barRaw : barRaw.slice(0, topX);
  const sortedLabels = barSeries.map((x) => x.label);
  const sortedValues = barSeries.map((x) => x.value);

  const canvas = $('wishbarShareChart');
  if (canvas) {
    if (chart) chart.destroy();
    if (sortedLabels.length) {
      const colors = palette(sortedLabels.length);
      canvas.parentElement.style.height = '420px';
      chart = new Chart(canvas.getContext('2d'), {
        type: 'bar',
        data: {
          labels: sortedLabels,
          datasets: [{
            label: '访问占比',
            data: sortedValues,
            backgroundColor: colors.map((c) => c + '44'),
            borderColor: colors,
            borderWidth: 1.2,
            borderRadius: 6,
            barThickness: 18,
          }],
        },
        options: {
          indexAxis: 'y',
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: 'nearest', intersect: true },
          layout: { padding: { top: 10, bottom: 10 } },
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: { label: (ctx) => `${ctx.label}: ${fmtPct(ctx.parsed.x)}` },
            },
          },
          scales: {
            x: {
              min: 0,
              max: Math.min(100, Math.ceil((Math.max(...sortedValues, 0) + 5) / 5) * 5 || 100),
              ticks: { callback: (v) => `${v}%` },
              grid: { color: 'rgba(15,23,42,.08)' },
            },
            y: { grid: { display: false } },
          },
        },
      });
    } else {
      chart = null;
    }
  }

  const barHint = sortedLabels.length
    ? `展示 Top${topX >= 999 ? '全部' : topX}，共 ${sortedLabels.length} 个一级来源（按占比倒序；已隐藏「(空)」）`
    : `当前周（${weekForBar}）暂无数据`;
  const hintEl = $('wishbarChartHint');
  if (hintEl) hintEl.textContent = barHint;

  // ----- 模块2：宣推贡献表格 -----
  const m2Week = $('wishbarModule2Week')?.value || latestWeek;
  const weekForM2 = fp.weeks.includes(m2Week) ? m2Week : latestWeek;
  const rowsM2 = buildWishbarContributionRows(fp, weekForM2);
  const tblBody = $('wishbarContributionTable')?.querySelector('tbody');
  if (tblBody) {
    tblBody.innerHTML = rowsM2
      .map((r) => `<tr>
        <td>${escapeHtml(r.l1)}</td>
        <td>${fmtPct(r.uvShare)}</td>
        <td>${fmtPct(r.drawShare)}</td>
        <td>${fmtPct(r.payUserShare)}</td>
        <td>${fmtPct(r.payAmountShare)}</td>
        <td class="mono">${fmtNum(r.uv)}</td>
        <td class="mono">${fmtNum(r.draw)}</td>
        <td class="mono">${fmtNum(r.payUser)}</td>
        <td class="mono">${fmtNum(r.payAmount)}</td>
        <td class="mono">${fmtPerKUsers(r.draw, r.uv)}</td>
        <td class="mono">${fmtPerKUsers(r.payUser, r.uv)}</td>
        <td class="mono">${fmtPerKAmount(r.payAmount, r.uv)}</td>
      </tr>`)
      .join('');
  }

  const insightEl = $('wishbarContributionInsight');
  const rawForReport = getRowsForTrendAggregate();
  if (insightEl) {
    if (rowsM2.length) {
      insightEl.innerHTML = buildWishbarWeeklyReportHtml(fp, weekForM2, rowsM2, rawForReport);
    } else {
      insightEl.innerHTML = '<p class="wishbarReport__muted">本周无一级来源明细，无法生成简报。</p>';
    }
  }

  // ----- 模块3：一级来源并集 + 本模块日期范围；折线（占比或指数） -----
  const trendWeeks = getTrendChartWeeks(fp);
  const fpTrend = sliceFpToWeekList(fp, trendWeeks.length ? trendWeeks : fp.weeks);
  const trendRows = getRowsForTrendAggregate();
  const unionSeries = aggregateTrendUnionSeries(trendRows, fpTrend, trendSelectedL1, new Set());

  const trendL1 = Array.from(trendSelectedL1);
  const yMode = $('wishbarTrendYMode')?.value || 'share';

  const volCanvas = $('wishbarVolumeChart');
  const volHintEl = $('wishbarVolumeHint');
  const metricDefs = [
    { key: 'uv', label: '访问占比', pct: unionSeries.uvPct, raw: unionSeries.uvRaw, fmtRaw: (v) => fmtNum(v) },
    { key: 'draw', label: '抽卡用户数占比', pct: unionSeries.drawPct, raw: unionSeries.drawRaw, fmtRaw: (v) => fmtNum(v) },
    { key: 'payUser', label: '付费抽卡用户数占比', pct: unionSeries.payUserPct, raw: unionSeries.payUserRaw, fmtRaw: (v) => fmtNum(v) },
    { key: 'payAmt', label: '付费金额占比', pct: unionSeries.payAmtPct, raw: unionSeries.payAmtRaw, fmtRaw: (v) => fmtNum(v) },
  ];

  if (volCanvas && fpTrend.weeks.length) {
    const colors = palette(4);
    const datasets = metricDefs.map((def, i) => {
      const series = yMode === 'index' ? toIndexFromFirst(def.pct) : def.pct;
      return {
        label: def.label,
        data: series,
        borderColor: colors[i],
        backgroundColor: colors[i] + '22',
        pointRadius: 0,
        pointHoverRadius: 4,
        tension: 0.25,
        fill: false,
        borderWidth: 2,
      };
    });

    volCanvas.parentElement.style.height = '400px';
    if (volumeChart) volumeChart.destroy();
    volumeChart = new Chart(volCanvas.getContext('2d'), {
      type: 'line',
      data: { labels: fpTrend.weeks, datasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        layout: { padding: { top: 10, bottom: 10 } },
        plugins: {
          legend: { position: 'bottom', align: 'center', labels: { usePointStyle: true, boxWidth: 10, padding: 12 } },
          tooltip: {
            callbacks: {
              label: (ctx) => {
                const def = metricDefs[ctx.datasetIndex];
                if (!def) return '';
                const idx = ctx.dataIndex;
                const pctVal = def.pct[idx];
                const raw = def.raw[idx];
                const rawStr = def.fmtRaw ? `，汇总 ${def.fmtRaw(raw)}` : '';
                if (yMode === 'index') {
                  const idxVal = ctx.parsed.y;
                  return `${def.label}: 指数 ${idxVal.toFixed(1)}（占比 ${fmtPct(pctVal)}${rawStr}）`;
                }
                return `${def.label}: ${fmtPct(ctx.parsed.y)}${rawStr}`;
              },
            },
          },
        },
        scales: {
          y: {
            min: yMode === 'index' ? undefined : 0,
            title: {
              display: true,
              text: yMode === 'index' ? '指数（所选区间首周=100）' : '占当周全量比例（%）',
            },
            ticks: {
              callback: (v) => (yMode === 'index' ? `${v}` : `${v}%`),
            },
            grid: { color: 'rgba(15,23,42,.08)' },
          },
          x: { grid: { display: false }, ticks: { maxRotation: 45, minRotation: 30, autoSkip: true, maxTicksLimit: 12 } },
        },
      },
    });

    if (volHintEl) {
      const unionDesc = !trendSelectedL1.size
        ? '未选一级来源（全量）'
        : `一级来源并集 ${trendL1.length} 项`;
      const rangeDesc = fpTrend.weeks.length
        ? `横轴 ${fpTrend.weeks[0]}～${fpTrend.weeks[fpTrend.weeks.length - 1]} 共 ${fpTrend.weeks.length} 周`
        : '';
      const modeHint = yMode === 'index'
        ? '纵轴为「指数」：看四条线谁相对首周涨得多，便于比形态；绝对占比请看悬停。'
        : '纵轴为「占当周全量%」：看所选来源合计在大盘中的体量；四条线可直接比高低。';
      volHintEl.textContent = `${unionDesc}。${rangeDesc}。${modeHint}`;
    }
  } else {
    if (volumeChart) { volumeChart.destroy(); volumeChart = null; }
    if (volHintEl) volHintEl.textContent = '暂无周度数据。';
  }

  // ----- 模块4：明细表 -----
  const tableEl = $('wishbarShareTable');
  if (tableEl) {
    const cols = fp.l1All.filter((c) => !isEmptyPrimarySource(c));
    let html = '<thead><tr><th>周</th>';
    for (const c of cols) html += `<th>${escapeHtml(c)}</th>`;
    html += '</tr></thead><tbody>';
    for (const w of fp.weeks) {
      const m = fp.l1SharesByWeek.get(w) || new Map();
      html += `<tr><td class="mono">${escapeHtml(w)}</td>`;
      for (const c of cols) html += `<td>${fmtPct(m.get(c) || 0)}</td>`;
      html += '</tr>';
    }
    html += '</tbody>';
    tableEl.innerHTML = html;
  }

  setStatus(`已载入 ${parsed.weeks.length} 个周，筛选后 ${fp.weeks.length} 周。当前周：${weekForBar}。`);
}

async function onFile(file) {
  setStatus('解析中…');
  try {
    const rows = await parseCsvFile(file);
    rawRows = rows;
    parsed = buildWeeklyShares(rows);
    trendInitialized = false;
    resetWeekDropdownCache();
    render();
  } catch (e) {
    parsed = null;
    setStatus(`解析失败：${e?.message || String(e)}`);
  }
}

async function loadLatestFromBoundFolder() {
  setStatus('读取绑定文件夹中…');
  try {
    const root = await getBoundDirHandle();
    if (!root) throw new Error('尚未绑定数据文件夹。');
    const perm = await root.queryPermission?.({ mode: 'read' });
    if (perm !== 'granted') {
      const req = await root.requestPermission?.({ mode: 'read' });
      if (req !== 'granted') throw new Error('未获得读取权限。');
    }
    const { handle: wishDir, name: subName } = await resolveWishSubdir(root);
    setStatus(`正在扫描「${subName}」下所有 CSV 并合并…`);
    const allRows = [];
    let fileCount = 0;
    for await (const entry of wishDir.values()) {
      if (!entry || entry.kind !== 'file') continue;
      if (!entry.name.toLowerCase().endsWith('.csv')) continue;
      const file = await entry.getFile();
      const rows = await parseCsvText(await file.text());
      for (const r of rows) allRows.push(r);
      fileCount += 1;
    }
    if (!fileCount) throw new Error('祈愿子文件夹下未找到任何 CSV 文件。');
    rawRows = allRows;
    parsed = buildWeeklyShares(allRows);
    trendInitialized = false;
    resetWeekDropdownCache();
    render();
    if (parsed.weeks.length > 0) {
      setStatus(`更新完成：已合并「${subName}」下 ${fileCount} 个 CSV。`);
    }
  } catch (e) {
    setStatus(`读取失败：${e?.message || String(e)}`);
  }
}

/** 本页可交互快照（单 HTML；数据极大时可能失败，请用汇总页 ZIP 或减小 CSV） */
async function exportWishbarPageSnapshot() {
  if (!rawRows.length) {
    alert('请先加载数据再导出本页快照。');
    return;
  }
  const btn = $('wishbarExportSnapshotBtn');
  const orig = btn?.textContent;
  if (btn) btn.textContent = '打包中…';
  try {
    const cssText = await fetch('./styles.css').then((r) => r.text()).catch(() => '');
    const jsText = await fetch(`./wishbar.js?${Date.now()}`).then((r) => r.text()).catch(() => '');
    const dataPayload = JSON.stringify({ wishbarRows: rawRows });
    const now = new Date();
    const ts = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    const pageHtml = document.querySelector('.appShell__content main')?.innerHTML || '';
    const headerHtml = document.querySelector('.appShell__content header')?.innerHTML || '';
    const snapCss =
      '.sidebar{display:none!important}.appShell{grid-template-columns:1fr!important}' +
      '.header__actions .file,.header__actions #wishbarBindFolderBtn,.header__actions #wishbarUpdateBtn,.header__actions #wishbarClearBtn,.header__actions #wishbarExportSnapshotBtn{display:none!important}' +
      '.header{position:static!important;top:auto!important}';
    const html = `<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>祈愿bar访问贡献快照 ${ts}</title>
<style>${cssText}</style>
<style>${snapCss}
.snap-banner{background:linear-gradient(135deg,#f0f4ff,#fdf4ff);border-bottom:1px solid rgba(79,124,255,.15);padding:10px 18px;font-size:12px;color:rgba(15,23,42,.7)}
.snap-banner strong{color:rgba(15,23,42,.9)}</style>
</head>
<body data-page="wishbar">
<div class="appShell">
<div class="appShell__content">
<div class="snap-banner"><span>\ud83d\udccb <strong>\u672c\u9875\u53ef\u4ea4\u4e92\u5feb\u7167</strong>\u3000\u751f\u6210\u65f6\u95f4\uff1a${ts}</span></div>
<header class="header">${headerHtml}</header>
<main class="container">${pageHtml}</main>
</div>
</div>
<script>window.__SNAPSHOT_DATA=${dataPayload};<\/script>
<script src="https://cdn.jsdelivr.net/npm/papaparse@5.4.1/papaparse.min.js"><\/script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"><\/script>
<script>${jsText}<\/script>
</body>
</html>`;
    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const dateStr = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
    a.download = `\u7948\u613f\u5ba3\u53d1\u5feb\u7167_${dateStr}.html`;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      URL.revokeObjectURL(url);
      a.remove();
    }, 1000);
    setStatus('本页快照已下载。');
  } catch (e) {
    setStatus(`导出失败：${e?.message || e}`);
  } finally {
    if (btn) btn.textContent = orig || '导出快照';
  }
}

function init() {
  const input = $('wishbarFileInput');
  input?.addEventListener('change', (e) => {
    const f = e.target.files?.[0];
    if (f) onFile(f);
  });

  document.addEventListener('dragover', (e) => e.preventDefault());
  document.addEventListener('drop', (e) => {
    e.preventDefault();
    const f = e.dataTransfer?.files?.[0];
    if (f) onFile(f);
  });

  $('wishbarWeekPick')?.addEventListener('change', render);
  $('wishbarModule1TopX')?.addEventListener('change', render);
  $('wishbarModule2Week')?.addEventListener('change', render);
  $('wishbarDateStart')?.addEventListener('change', render);
  $('wishbarDateEnd')?.addEventListener('change', render);
  $('wishbarTrendWeekStart')?.addEventListener('change', render);
  $('wishbarTrendWeekEnd')?.addEventListener('change', render);
  $('wishbarTrendYMode')?.addEventListener('change', render);

  function renderGlobalL1Summary() {
    const el = $('wishbarGlobalL1Summary');
    if (el) el.textContent = summarizeSelection(globalL1Set);
  }
  function renderGlobalL2Summary() {
    const el = $('wishbarGlobalL2Summary');
    if (el) el.textContent = summarizeSelection(globalL2Set);
  }

  $('wishbarGlobalL1Btn')?.addEventListener('click', () => {
    if (!parsed) return;
    const items = parsed.l1All;
    openFilterPopover({
      anchorEl: $('wishbarGlobalL1Btn'),
      title: '一级来源筛选（全局）',
      items,
      selectedSet: globalL1Set,
      onApply: (next) => {
        globalL1Set = next;
        renderGlobalL1Summary();
        render();
      },
    });
  });

  $('wishbarGlobalL2Btn')?.addEventListener('click', () => {
    if (!parsed) return;
    const items = parsed.l2All;
    openFilterPopover({
      anchorEl: $('wishbarGlobalL2Btn'),
      title: '二级来源筛选（全局）',
      items,
      selectedSet: globalL2Set,
      onApply: (next) => {
        globalL2Set = next;
        renderGlobalL2Summary();
        render();
      },
    });
  });

  $('wishbarTrendL1Btn')?.addEventListener('click', () => {
    if (!parsed) return;
    const fp = getFilteredParsed() || parsed;
    const order = buildSourceOrderByLatest(fp);
    const l1NE = fp.l1All.filter((s) => !isEmptyPrimarySource(s));
    const items = order.length < l1NE.length ? [...order, '其他'] : order.slice();
    openFilterPopover({
      anchorEl: $('wishbarTrendL1Btn'),
      title: '一级来源（模块三趋势）',
      items,
      selectedSet: trendSelectedL1,
      onApply: (next) => {
        trendSelectedL1 = next;
        const s = $('wishbarTrendL1Summary');
        if (s) s.textContent = summarizeSelection(next);
        render();
      },
    });
  });

  $('wishbarBindFolderBtn')?.addEventListener('click', async () => {
    setStatus('等待选择文件夹…');
    try {
      await pickAndBindFolder();
      setStatus('已绑定。点击「读取祈愿文件夹全部CSV并更新」。');
    } catch (e) {
      const msg = e?.message || String(e);
      setStatus(`绑定失败：${msg}`);
      if (msg.includes('iframe')) {
        try { window.open('./wishbar.html', '_blank', 'noopener,noreferrer'); } catch (_) {}
      }
    }
  });

  $('wishbarUpdateBtn')?.addEventListener('click', loadLatestFromBoundFolder);

  $('wishbarClearBtn')?.addEventListener('click', () => {
    input.value = '';
    rawRows = [];
    parsed = null;
    globalL1Set = new Set();
    globalL2Set = new Set();
    trendSelectedL1 = new Set();
    trendInitialized = false;
    resetWeekDropdownCache();
    const ds = $('wishbarDateStart');
    const de = $('wishbarDateEnd');
    if (ds) ds.innerHTML = '';
    if (de) de.innerHTML = '';
    syncModuleWeekDropdowns({ weeks: [] });
    render();
    setStatus('已清空。请选择/拖入 CSV 文件开始分析。');
    renderGlobalL1Summary();
    renderGlobalL2Summary();
  });

  $('wishbarExportSnapshotBtn')?.addEventListener('click', () => {
    exportWishbarPageSnapshot();
  });

  if (window.__SNAPSHOT_DATA?.wishbarRows) {
    rawRows = window.__SNAPSHOT_DATA.wishbarRows;
    parsed = buildWeeklyShares(rawRows);
    trendInitialized = false;
    resetWeekDropdownCache();
    render();
    setStatus(`快照模式：已加载 ${rawRows.length} 行。`);
  } else {
    getBoundDirHandle().then(async (h) => {
      if (!h) return;
      try {
        const perm = await h.queryPermission?.({ mode: 'read' });
        if (perm === 'granted') await loadLatestFromBoundFolder();
        else setStatus('检测到已绑定数据文件夹，请点击「读取祈愿文件夹全部CSV并更新」授权读取。');
      } catch (_) {
        setStatus('检测到已绑定数据文件夹，请点击「读取祈愿文件夹全部CSV并更新」。');
      }
    });
  }
}

init();
