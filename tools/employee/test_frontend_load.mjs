import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';

const root = new URL('../../', import.meta.url);

function extractFunction(source, signature) {
  const start = source.indexOf(signature);
  if (start < 0) throw new Error(`function not found: ${signature}`);
  const tail = source.slice(start);
  const match = tail.match(/^[\s\S]*?^}\r?$/m);
  if (!match) throw new Error(`function end not found: ${signature}`);
  return match[0];
}

function loadFrontend(file) {
  const source = fs.readFileSync(new URL(file, root), 'utf8');
  const pendingFetches = [];
  const alerts = [];
  const loadingStates = [];
  const noOp = () => {};
  const context = vm.createContext({
    console: { error() {} },
    currentCompany: '会社A',
    dataLoadSeq: 0,
    dataLoadOk: false,
    allNippos: [],
    allMembers: [],
    allGenbaMaster: [],
    allJobsites: [],
    isSubmitting: false,
    GAS_URL: 'https://example.invalid/exec',
    encodeURIComponent,
    fetch() {
      return new Promise(resolve => pendingFetches.push(resolve));
    },
    parseRows: rows => rows,
    generateGhosts: () => [],
    setLoading: value => loadingStates.push(value),
    showAlert: message => alerts.push(message),
    initMemberUI: noOp,
    initGenbaSelects: noOp,
    renderList: noOp,
    updateAvailability: noOp,
    updateJimu: noOp,
    updateGenbaFilter: noOp,
    refreshGmTab: noOp,
    updateLocationSuggestions: noOp,
    refreshVehicleSelect: noOp,
    checkUsername: noOp,
    document: {
      getElementById: () => ({ disabled: false }),
    },
  });
  vm.runInContext(extractFunction(source, 'async function loadData(){'), context);
  vm.runInContext(extractFunction(source, 'async function submitNippo(){'), context);
  return { context, pendingFetches, alerts, loadingStates };
}

for (const file of ['index.html', 'admin.html']) {
  test(`${file} ignores a late response from the previously selected company`, async () => {
    const app = loadFrontend(file);

    const first = app.context.loadData();
    app.context.currentCompany = '会社B';
    const second = app.context.loadData();

    app.pendingFetches[1]({
      json: async () => ({
        status: 'ok',
        rows: [{ 会社: '会社B', ID: 'B' }],
        members: [{ name: 'B社員' }],
        genbaMaster: [],
        jobsites: [],
      }),
    });
    await second;
    app.pendingFetches[0]({
      json: async () => ({
        status: 'ok',
        rows: [{ 会社: '会社A', ID: 'A' }],
        members: [{ name: 'A社員' }],
        genbaMaster: [],
        jobsites: [],
      }),
    });
    await first;

    assert.deepEqual(app.context.allNippos, [{ 会社: '会社B', ID: 'B' }]);
    assert.deepEqual(app.context.allMembers, [{ name: 'B社員' }]);
    assert.equal(app.context.dataLoadOk, true);
  });

  test(`${file} clears stale data and blocks registration after a load error`, async () => {
    const app = loadFrontend(file);
    app.context.allNippos = [{ 会社: '会社A', ID: 'OLD' }];
    app.context.allMembers = [{ name: '古い社員' }];

    const load = app.context.loadData();
    app.pendingFetches[0]({
      json: async () => ({ status: 'error', message: 'busy' }),
    });
    await load;

    assert.equal(app.context.allNippos.length, 0);
    assert.equal(app.context.allMembers.length, 0);
    assert.equal(app.context.dataLoadOk, false);
    assert.match(app.alerts[0], /読込エラー/);

    await app.context.submitNippo();
    assert.match(app.alerts.at(-1), /最新データを読み込めていない/);
  });
}
