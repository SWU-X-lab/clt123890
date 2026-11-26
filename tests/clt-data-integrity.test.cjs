// 回归测试：以无网络 DOM / Supabase / XLSX 存根执行默认入口加载的 clt.js。
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const source = fs.readFileSync(path.join(__dirname, '..', 'clt.js'), 'utf8');

function harness({ group = '1', code = 'Study_ABC123', failFullRow = false } = {}) {
    const elements = new Map();
    const windowListeners = {};
    const documentListeners = {};
    const timers = new Map();
    let timer = 0;
    let time = 100;
    let focused = true;
    let initialize = 0;
    const writes = [];
    const sheets = {};
    const noop = () => {};
    const canvasContext = new Proxy({}, { get: (t, key) => t[key] || noop });
    function element(id) {
        if (!elements.has(id)) elements.set(id, {
            listeners: {}, classList: { add: noop, remove: noop },
            value: id === 'subGroup' ? group : id === 'participantCode' ? code : '',
            addEventListener(type, fn) { this.listeners[type] = fn; },
            reportValidity: () => true,
            getContext: () => canvasContext,
            width: 1024, height: 768
        });
        return elements.get(id);
    }
    const document = {
        getElementById: element, documentElement: {}, hidden: false,
        hasFocus: () => focused, exitFullscreen: noop,
        addEventListener(type, listener) { documentListeners[type] = listener; }
    };
    const window = {
        addEventListener(type, listener) { windowListeners[type] = listener; },
        supabase: { createClient: () => {
            initialize++;
            return { from: () => ({ insert: async rows => {
                writes.push(JSON.parse(JSON.stringify(rows)));
                return { error: failFullRow && writes.length === 1 ? new Error('测试列缺失') : null };
            } }) };
        } }
    };
    const XLSX = { utils: {
        json_to_sheet(data, options) { return { data, header: options.header }; }
    } };
    const fixture = vm.runInNewContext(source + `
        ;({
            buildKRow, kSheetFromRows, participantSeed, finishExperiment, startFormal, runTrial,
            setP(v) { p = v; }, getP() { return p; }, session: SESSION_ID,
            setData(rows) { stimData = rows; }, rows() { return stimData; },
            stopDraw() { drawBackground = () => {}; drawFixation = () => {}; },
            mockExport(fn) { downloadXLSX = fn; },
            setPhase(v) { experimentPhase = v; }, phase() { return experimentPhase; },
            sequence() { startBlock = () => {}; return () => { startFormal(); return allFormalTrials; }; },
            seedInternals(n) { seedRandom(n); },
            setActive(v) {
                trialPlan = { setSize: 4, labelMap: [2, 0, 3, 1], changeIdx: 2, probeColors: [0, 1, 8, 3] };
                applyPlanToGlobals({ ...trialPlan, memColors: [0, 1, 2, 3], newColor: 8, correctLabel: 1, trialLabels: [0, 1, 2, 3] });
                trialData = v;
                experimentPhase = 'trial_response';
            },
            setTimer(id) { responseTimer = id; }
        })`, {
            document, window, XLSX,
            performance: { now: () => time },
            setTimeout(fn) { timers.set(++timer, fn); return timer; },
            clearTimeout(id) { timers.delete(id); },
            requestAnimationFrame(fn) { timers.set(++timer, fn); },
            console: { log: noop, error: noop, warn: noop },
            alert: noop, crypto: { randomUUID: () => 'mock-run-id' }
        });
    fixture.stopDraw();
    fixture.mockExport(() => { sheets.downloads = (sheets.downloads || 0) + 1; });
    return { fixture, document, elements, getInitialize: () => initialize, writes, sheets, timers,
        start: () => element('start-btn').listeners.click(),
        executeKey: key => windowListeners.keydown({ key, code: key === ' ' ? 'Space' : '', repeat: false, preventDefault: noop }),
        hide: () => { document.hidden = true; focused = false; documentListeners.visibilitychange(); },
        show: () => { document.hidden = false; focused = true; },
        blur: () => { focused = false; windowListeners.blur(); }
    };
}
const row = (accuracy, rt, extra = {}) => ({ setSize: 4, accuracy, rt, CountedInK: 1, TrialStatus: 'valid', ...extra });

test('空条件与全超时的 K 统计缺失值不落成 Excel 0；真实 0 保留', () => {
    const { fixture: f } = harness();
    const empty = f.buildKRow(4, []);
    const timeoutOnly = f.buildKRow(4, [row(0, null, { CountedInK: 0, timeout: true, TrialStatus: 'timeout' })]);
    const zero = f.buildKRow(4, [row(0, 1)]);
    const data = f.kSheetFromRows([empty, timeoutOnly, zero]).data;

    assert.equal(Object.hasOwn(data[0], 'Accuracy_TimeoutErr'), false);
    assert.equal(Object.hasOwn(data[1], 'Accuracy_Exclude'), false);
    assert.equal(Object.hasOwn(data[1], 'K_Exclude'), false);
    assert.equal(data[1].Accuracy_TimeoutErr, 0);
    assert.equal(data[1].K_TimeoutErr, -1.3333);
    assert.equal(data[2].Accuracy_Exclude, 0);
    assert.equal(data[2].K_Exclude, -1.3333);
});

test('RT 汇总只使用正确且有效试次；中断不是超时且不进入两种 K 分母', () => {
    const { fixture: f } = harness();
    const c = f.buildKRow(4, [
        row(1, 0.4), row(0, 1.6),
        row(null, null, { CountedInK: 0, TrialStatus: 'interrupted' }),
        row(0, null, { CountedInK: 0, TrialStatus: 'timeout', timeout: true })
    ]);
    assert.equal(c.MeanRT, 0.4);
    assert.equal(c.Accuracy_Exclude, 0.5);
    assert.equal(c.Accuracy_TimeoutErr, 0.3333);
    assert.equal(c.TimeoutTrials, 1);
    assert.equal(c.InterruptedTrials, 1);
});

test('组别 0 路径既不初始化也不写入 Supabase，完成后本地导出仍触发', async () => {
    const h = harness({ group: '0' });
    await h.start();
    h.fixture.setData([row(1, 0.4)]);
    await h.fixture.finishExperiment();
    assert.equal(h.fixture.getP().subjectID, '1111');
    assert.equal(h.getInitialize(), 0);
    assert.equal(h.writes.length, 0);
    assert.equal(h.sheets.downloads, 1);
});

test('正式组可以使用同一匿名编号且顺序可复现；保留大小、80 试次和确定种子', async () => {
    const a = harness();
    const b = harness();
    await a.start();
    await b.start();
    const makeA = a.fixture.sequence();
    const makeB = b.fixture.sequence();
    a.fixture.seedInternals(101);
    b.fixture.seedInternals(123);
    const x = JSON.parse(JSON.stringify(makeA()));
    const y = JSON.parse(JSON.stringify(makeB()));
    assert.equal(a.fixture.getP().subjectID, 'Study_ABC123');
    assert.equal(a.fixture.getP().rndSeed, b.fixture.getP().rndSeed);
    assert.deepEqual(x, y);
    assert.equal(x.length, 80);
    assert.equal(x.filter(r => r.setSize === 4).length, 40);
    assert.equal(x.filter(r => r.setSize === 6).length, 40);
    assert.equal(a.getInitialize(), 1);
});

test('未分配编号时，页面预览与实际提交保持同一个自动编号', async () => {
    const h = harness({ code: '' });
    const expected = h.elements.get('subjectID').value;
    assert.match(expected, /^1_[0-9]{3}_[a-z0-9]{3}$/);
    await h.start();
    assert.equal(h.fixture.getP().subjectID, expected);
});

test('正式组保留旧表基础字段兼容，但回退时明显提示，应保留本地完整文件', async () => {
    const h = harness({ failFullRow: true });
    await h.start();
    h.fixture.setData([row(1, 0.4, { responseColorIdx: 8 })]);
    await h.fixture.finishExperiment();
    assert.equal(h.writes.length, 2);
    assert.equal(h.writes[0][0].responseColorIdx, 8);
    assert.equal(Object.hasOwn(h.writes[1][0], 'responseColorIdx'), false);
    assert.match(h.elements.get('finished-text').innerText, /仅保存了基础字段/);
    assert.equal(h.sheets.downloads, 1);
});

test('数字 1 作答写入对应探测色，而非标签槽原色；重复输入不再记第二条', () => {
    const h = harness();
    h.fixture.setActive({ setSize: 4 });
    h.executeKey('1');
    h.executeKey('1');
    const rows = h.fixture.rows();
    assert.equal(rows.length, 1);
    assert.equal(rows[0].responseKey, '1');
    assert.equal(rows[0].responseColorIdx, 8);
    assert.equal(rows[0].correctColorIdx, 8);
    assert.equal(rows[0].accuracy, 1);
});

test('正式作答中隐藏标签页会取消计时，将中断另记一次，不误算超时', () => {
    const h = harness();
    h.fixture.setActive({ setSize: 4, TrialGlobal: 1 });
    h.timers.set(101, () => { throw new Error('旧计时不应执行'); });
    h.fixture.setTimer(101);
    h.hide();
    h.blur(); // visibilitychange 和 blur 在同一试次均可能触发
    assert.equal(h.timers.has(101), false);
    assert.equal(h.fixture.phase(), 'interruption_pause');
    assert.equal(h.fixture.rows().length, 1);
    assert.equal(h.fixture.rows()[0].TrialStatus, 'interrupted');
    assert.equal(h.fixture.rows()[0].CountedInK, 0);
    assert.equal(h.fixture.rows()[0].timeout, false);
    assert.equal(h.fixture.rows()[0].InterruptionReason, 'visibility_hidden');
});
test('两个入口都对齐匿名输入与新构建版本，仍加载同一脚本', () => {
    for (const file of ['index.html', 'clt_123890.html']) {
        const html = fs.readFileSync(path.join(__dirname, '..', file), 'utf8');
        assert.match(html, /id="participantCode"/);
        assert.match(html, /clt\.js\?v=20261126-123890-qa1/);
        assert.match(html, /CLT_PAGE_BUILD = '20261126-123890-qa1'/);
    }
});
