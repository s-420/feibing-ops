/**
 * @对象    门店
 * @动作    批量写
 * @风险    高
 * @来源    store-data-extractor/PiLiangShuaXinMenDianKeFuYiDuiYi.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

/**
 * “门店客服批量刷”一店一客服安全执行脚本。
 * 默认只预演；正式执行需 --execute + CUSTOMER_SERVICE_WRITE_CONFIRM=门店客服批量刷193家。
 */
require('dotenv').config();

const fs = require('fs');
const path = require('path');
const axios = require('axios');
const xlsx = require('xlsx');

const config = {
  sellerId: 'wwa9c5a585540b115b',
  // 沪上阿姨客服用手机号直接标识（userID 即手机号），无需 phone→userID。
  phoneAsUserId: true,
  // 客服通讯录归属 ID；不能用门店资产 sellerId，否则手机号会查不到 userID，写接口会“假成功”。
  contactDirectorySellerId: 'wpV3cNDAAA3ZRQDQZt0UsR9wORkx0thg',
  sourcePath: 'D:/download/飞冰每日异常门店统计.xlsx',
  mappingSheet: '门店客服批量刷',
  masterSheet: '主数据',
  expectedCount: 193,
  outputDir: path.join(__dirname, '../output'),
  readConcurrency: 3,
  requestDelayMs: 150,
  timeoutMs: 30000,
};

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const stamp = () => new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14);

function readMappings(sourcePath = config.sourcePath) {
  if (!fs.existsSync(sourcePath)) throw new Error(`源文件不存在: ${sourcePath}`);
  const workbook = xlsx.readFile(sourcePath);
  const mappingSheet = workbook.Sheets[config.mappingSheet];
  const masterSheet = workbook.Sheets[config.masterSheet];
  if (!mappingSheet || !masterSheet) throw new Error('缺少“门店客服批量刷”或“主数据”工作表');

  const mappingRows = xlsx.utils.sheet_to_json(mappingSheet, { defval: '' });
  const masterRows = xlsx.utils.sheet_to_json(masterSheet, { header: 1, defval: '' });
  const masterByName = new Map();
  for (let index = 1; index < masterRows.length; index += 1) {
    const storeName = String(masterRows[index][1] || '').trim(); // 主数据 B 列
    if (!storeName) continue;
    const items = masterByName.get(storeName) || [];
    items.push({ masterRow: index + 1, shopId: String(masterRows[index][17] || '').trim() }); // 主数据 R 列
    masterByName.set(storeName, items);
  }

  const mappings = mappingRows.map((row, index) => {
    const storeName = String(row['门店名称'] || '').trim();
    const matches = masterByName.get(storeName) || [];
    return {
      sourceRow: index + 2,
      storeCode: String(row['门店编码'] || '').trim(),
      storeName,
      customerAccount: String(row['客服账号'] || '').trim(),
      masterMatches: matches,
      shopId: matches.length === 1 ? matches[0].shopId : '',
    };
  });
  validateMappings(mappings);
  return mappings.map(({ masterMatches, ...mapping }) => mapping);
}

function validateMappings(mappings) {
  if (mappings.length !== config.expectedCount) {
    throw new Error(`映射数量必须为 ${config.expectedCount}，当前为 ${mappings.length}`);
  }
  const names = new Set();
  const ids = new Set();
  for (const item of mappings) {
    if (!item.storeName || !/^\d{11}$/.test(item.storeCode)) {
      throw new Error(`第 ${item.sourceRow} 行门店编码或名称不合法`);
    }
    if (!/^\d{11}$/.test(item.customerAccount)) {
      throw new Error(`第 ${item.sourceRow} 行客服账号不是 11 位数字: ${item.customerAccount || '(空)'}`);
    }
    if (item.masterMatches.length !== 1) {
      throw new Error(`第 ${item.sourceRow} 行门店“${item.storeName}”在主数据匹配 ${item.masterMatches.length} 条`);
    }
    if (!/^[a-f0-9]{24}$/i.test(item.shopId)) {
      throw new Error(`第 ${item.sourceRow} 行店铺ID不合法: ${item.shopId || '(空)'}`);
    }
    if (names.has(item.storeName) || ids.has(item.shopId)) {
      throw new Error(`发现重复门店: ${item.storeName} / ${item.shopId}`);
    }
    names.add(item.storeName);
    ids.add(item.shopId);
  }
}

function authorization() {
  const value = String(process.env.VINCI_AUTHORIZATION || '').trim();
  if (!value) throw new Error('缺少 VINCI_AUTHORIZATION');
  return /^bearer\s+/i.test(value) ? value : `bearer ${value}`;
}

function createClient() {
  const headers = {
    accept: 'application/json',
    authorization: authorization(),
    'content-type': 'application/json;charset=UTF-8',
    origin: 'https://connect.feibing.tech',
    referer: 'https://connect.feibing.tech/',
    'user-agent': 'Mozilla/5.0',
  };
  if (process.env.VINCI_COOKIE) headers.cookie = process.env.VINCI_COOKIE;
  return axios.create({ headers, timeout: config.timeoutMs, validateStatus: () => true });
}

function assertSuccess(response, action) {
  if (response.status === 401 || response.status === 403) throw new Error(`${action}: 凭据无效（HTTP ${response.status}）`);
  if (response.status < 200 || response.status >= 300) throw new Error(`${action}: HTTP ${response.status}`);
  if (response.data?.code !== 0) throw new Error(`${action}: ${response.data?.message || `code=${response.data?.code}`}`);
  return response.data;
}

async function readWithRetry(action, request, attempts = 3) {
  let lastError;
  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      const response = await request();
      if ((response.status === 429 || response.status >= 500) && attempt < attempts) throw new Error(`HTTP ${response.status}`);
      return response;
    } catch (error) {
      lastError = error;
      if (attempt < attempts) await delay(attempt * 800);
    }
  }
  throw new Error(`${action}连续 ${attempts} 次失败: ${lastError?.message}`);
}

async function mapConcurrent(items, concurrency, worker, onDone) {
  const results = new Array(items.length);
  let cursor = 0;
  async function run() {
    while (cursor < items.length) {
      const index = cursor;
      cursor += 1;
      results[index] = await worker(items[index], index);
      if (onDone) onDone(index, results[index]);
    }
  }
  await Promise.all(Array.from({ length: Math.min(concurrency, items.length) }, run));
  return results;
}

async function resolveCustomer(client, account) {
  if (config.phoneAsUserId) {
    return { userId: account, lookupStatus: 'phone-as-userid' };
  }
  const response = await readWithRetry(`查询客服 ${account}`, () => client.post(
    'https://vinci-api.feibing.tech/xc/v1/cps/token/contact/search',
    { queryWord: account, sellerId: config.contactDirectorySellerId },
  ));
  const data = assertSuccess(response, `查询客服 ${account}`);
  const userIds = Array.isArray(data?.data?.user?.userid) ? data.data.user.userid.filter(Boolean).map(String) : [];
  if (userIds.length > 1) throw new Error(`客服账号 ${account} 匹配到 ${userIds.length} 个 userID`);
  if (userIds.length === 0) throw new Error(`客服账号 ${account} 未找到对应 userID，拒绝提交手机号`);
  return { userId: userIds[0], lookupStatus: userIds[0] === account ? 'confirmed-same' : 'confirmed-mapped' };
}

async function getCurrentOwners(client, shopId) {
  const response = await readWithRetry(`查询门店 ${shopId}`, () => client.get(
    `https://vinci-api.feibing.tech/sc/v1/sellers/${config.sellerId}/places`,
    { params: { current: 1, pageSize: 100, pageNum: 1, shopId, catalogId: '', types: 'DESK,SINGLE,CHANNEL,GROUP' } },
  ));
  const data = assertSuccess(response, `查询门店 ${shopId}`);
  const places = Array.isArray(data.data) ? data.data : [];
  const owners = new Set();
  let missingContact = 0;
  let parseErrors = 0;
  for (const place of places) {
    if (!place?.metadata?.contact) { missingContact += 1; continue; }
    try {
      const contact = JSON.parse(place.metadata.contact);
      if (Array.isArray(contact?.owners)) contact.owners.forEach((owner) => owner && owners.add(String(owner)));
    } catch (_) { parseErrors += 1; }
  }
  return { placeCount: places.length, owners: [...owners], missingContact, parseErrors };
}

async function updateContact(client, shopId, owners) {
  const response = await client.post(
    `https://vinci-api.feibing.tech/sc/v1/sellers/${config.sellerId}/places/${shopId}/batch/update/contact`,
    { batchUpdateType: 0, shopIds: [shopId], customerService: JSON.stringify({ nickName: owners[0] }), owners },
  );
  return assertSuccess(response, `更新门店 ${shopId}`);
}

function writeJson(prefix, value) {
  fs.mkdirSync(config.outputDir, { recursive: true });
  const filePath = path.join(config.outputDir, `${prefix}_${stamp()}.json`);
  fs.writeFileSync(filePath, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
  return filePath;
}

async function buildPreflight(client, mappings) {
  const uniqueAccounts = [...new Set(mappings.map((item) => item.customerAccount))];
  console.log(`开始验证 ${uniqueAccounts.length} 个唯一客服账号...`);
  let accountDone = 0;
  const resolutions = await mapConcurrent(uniqueAccounts, config.readConcurrency, async (account) => {
    const resolution = await resolveCustomer(client, account);
    await delay(config.requestDelayMs);
    return resolution;
  }, () => {
    accountDone += 1;
    if (accountDone % 20 === 0 || accountDone === uniqueAccounts.length) console.log(`客服验证 ${accountDone}/${uniqueAccounts.length}`);
  });
  const accountMap = new Map(uniqueAccounts.map((account, index) => [account, resolutions[index]]));
  const unresolvedAccounts = resolutions.filter((item) => item.lookupStatus === 'not-found-use-account').length;
  console.log(`客服账号验证完成：接口确认 ${uniqueAccounts.length - unresolvedAccounts}，未返回 userID ${unresolvedAccounts}（保留表内账号）`);

  console.log(`开始查询 ${mappings.length} 家门店当前客服...`);
  let storeDone = 0;
  const records = await mapConcurrent(mappings, config.readConcurrency, async (mapping) => {
    const current = await getCurrentOwners(client, mapping.shopId);
    if (current.placeCount === 0) throw new Error(`门店 ${mapping.storeName} 未查到桌台`);
    if (current.missingContact > 0 || current.parseErrors > 0) {
      throw new Error(`门店 ${mapping.storeName} 存在缺失/损坏的 contact 配置`);
    }
    await delay(config.requestDelayMs);
    const resolution = accountMap.get(mapping.customerAccount);
    return { ...mapping, targetUserId: resolution.userId, customerLookupStatus: resolution.lookupStatus, ...current };
  }, () => {
    storeDone += 1;
    if (storeDone % 20 === 0 || storeDone === mappings.length) console.log(`门店预检 ${storeDone}/${mappings.length}`);
  });
  return records;
}

async function execute(client, records) {
  const backupPath = writeJson('备份_门店客服批量刷一店一客服', {
    createdAt: new Date().toISOString(), sellerId: config.sellerId, sourcePath: config.sourcePath, records,
  });
  console.log(`写前备份: ${backupPath}`);
  const results = [];
  for (let index = 0; index < records.length; index += 1) {
    const record = records[index];
    if (record.owners.length === 1 && record.owners[0] === record.targetUserId) {
      results.push({ shopId: record.shopId, storeName: record.storeName, status: 'already-correct' });
      continue;
    }
    try {
      await updateContact(client, record.shopId, [record.targetUserId]);
      results.push({ shopId: record.shopId, storeName: record.storeName, status: 'accepted' });
    } catch (error) {
      results.push({ shopId: record.shopId, storeName: record.storeName, status: 'failed', reason: error.message });
    }
    if ((index + 1) % 20 === 0 || index + 1 === records.length) console.log(`写入进度 ${index + 1}/${records.length}`);
    await delay(300);
  }
  return { backupPath, resultPath: writeJson('结果_门店客服批量刷一店一客服', { backupPath, results }), results };
}

async function rollback(client, backupPath) {
  const backup = JSON.parse(fs.readFileSync(backupPath, 'utf8'));
  if (backup.sellerId !== config.sellerId || !Array.isArray(backup.records)) throw new Error('备份文件不合法');
  const results = [];
  for (let index = 0; index < backup.records.length; index += 1) {
    const record = backup.records[index];
    if (!record.owners.length) { results.push({ storeName: record.storeName, status: 'skipped-empty' }); continue; }
    try {
      await updateContact(client, record.shopId, record.owners);
      results.push({ storeName: record.storeName, status: 'accepted' });
    } catch (error) { results.push({ storeName: record.storeName, status: 'failed', reason: error.message }); }
    await delay(300);
  }
  return writeJson('结果_门店客服批量刷一店一客服_回滚', { backupPath, results });
}

function parseArgs(argv) {
  const rollbackArg = argv.find((arg) => arg.startsWith('--rollback='));
  const limitArg = argv.find((arg) => arg.startsWith('--limit='));
  const limit = limitArg ? Number(limitArg.slice(8)) : 0;
  if (!Number.isInteger(limit) || limit < 0) throw new Error('--limit 必须是非负整数');
  return { checkInput: argv.includes('--check-input'), execute: argv.includes('--execute'), rollbackPath: rollbackArg ? rollbackArg.slice(11) : '', limit };
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const mappings = readMappings();
  console.log(`输入校验通过：${mappings.length} 家门店，一店一客服，${new Set(mappings.map((item) => item.customerAccount)).size} 个唯一客服账号`);
  if (args.checkInput) return;
  const client = createClient();
  if (args.rollbackPath) { console.log(`回滚结果: ${await rollback(client, path.resolve(args.rollbackPath))}`); return; }
  const records = await buildPreflight(client, mappings);
  const alreadyCorrect = records.filter((record) => record.owners.length === 1 && record.owners[0] === record.targetUserId).length;
  const preflightPath = writeJson('预演_门店客服批量刷一店一客服', {
    createdAt: new Date().toISOString(), sellerId: config.sellerId, sourcePath: config.sourcePath,
    total: records.length, alreadyCorrect, needChange: records.length - alreadyCorrect, records,
  });
  console.log(`预演完成：已正确 ${alreadyCorrect}，待修改 ${records.length - alreadyCorrect}，报告: ${preflightPath}`);
  if (!args.execute) { console.log('预演模式，未调用写入接口。'); return; }
  if (process.env.CUSTOMER_SERVICE_WRITE_CONFIRM !== '门店客服批量刷193家') throw new Error('缺少正式写入二次确认');
  const executionRecords = args.limit > 0
    ? records.filter((record) => !(record.owners.length === 1 && record.owners[0] === record.targetUserId)).slice(0, args.limit)
    : records;
  const result = await execute(client, executionRecords);
  console.log(`执行结果: ${result.resultPath}`);
}

if (require.main === module) main().catch((error) => { console.error(`执行失败: ${error.message}`); process.exit(1); });

module.exports = { config, readMappings, validateMappings, parseArgs, mapConcurrent };
