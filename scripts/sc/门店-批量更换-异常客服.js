/**
 * @对象    门店
 * @动作    批量写
 * @风险    高
 * @来源    store-data-extractor/GengHuanBaYueSanShiYiRiYiChangMenDianKeFu.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

/**
 * 8 月 31 日橙色异常门店批量更换客服（安全专用版）
 *
 * 默认行为：预演，只查询目标客服和 17 家门店当前客服，不写入。
 * 正式执行：必须同时传 --execute，并设置 CUSTOMER_SERVICE_WRITE_CONFIRM=8月31日17家门店。
 * 回滚：传 --rollback=output/备份_8月31日异常门店换客服_时间戳.json。
 */

require('dotenv').config();

const fs = require('fs');
const path = require('path');
const axios = require('axios');
const xlsx = require('xlsx');

const config = {
  sellerId: 'wwa9c5a585540b115b',
  inputPath: path.join(__dirname, '../input/8月31日橙色门店更换客服.xlsx'),
  outputDir: path.join(__dirname, '../output'),
  targetAccount: '19538549785',
  targetName: '总部福利官-青苹果',
  expectedStoreCount: 17,
  requestDelayMs: 300,
  timeoutMs: 30000,
};

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function timestamp() {
  return new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14);
}

function ensureOutputDir() {
  fs.mkdirSync(config.outputDir, { recursive: true });
}

function readStores(inputPath = config.inputPath) {
  if (!fs.existsSync(inputPath)) {
    throw new Error(`输入文件不存在: ${inputPath}`);
  }
  const workbook = xlsx.readFile(inputPath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = xlsx.utils.sheet_to_json(sheet, { defval: '' });
  const stores = rows.map((row, index) => ({
    rowNumber: index + 2,
    shopId: String(row['店铺ID'] || '').trim(),
    shopName: String(row['门店名称'] || '').trim(),
    targetAccount: String(row['目标客服账号'] || '').trim(),
    targetName: String(row['目标客服名称'] || '').trim(),
    sourceRow: Number(row['来源行']),
  }));
  validateStores(stores);
  return stores;
}

function validateStores(stores) {
  if (stores.length !== config.expectedStoreCount) {
    throw new Error(`门店数量必须为 ${config.expectedStoreCount}，当前为 ${stores.length}`);
  }
  const ids = new Set();
  for (const store of stores) {
    if (!/^[a-f0-9]{24}$/i.test(store.shopId)) {
      throw new Error(`第 ${store.rowNumber} 行店铺ID不合法: ${store.shopId || '(空)'}`);
    }
    if (!store.shopName) {
      throw new Error(`第 ${store.rowNumber} 行门店名称为空`);
    }
    if (store.targetAccount !== config.targetAccount || store.targetName !== config.targetName) {
      throw new Error(`第 ${store.rowNumber} 行目标客服不一致，拒绝执行`);
    }
    if (!Number.isInteger(store.sourceRow) || store.sourceRow < 5 || store.sourceRow > 21) {
      throw new Error(`第 ${store.rowNumber} 行来源行不在 5~21: ${store.sourceRow}`);
    }
    if (ids.has(store.shopId)) {
      throw new Error(`发现重复店铺ID: ${store.shopId}`);
    }
    ids.add(store.shopId);
  }
}

function getAuthorization() {
  const authorization = String(process.env.VINCI_AUTHORIZATION || '').trim();
  if (!authorization) {
    throw new Error('缺少 VINCI_AUTHORIZATION，请先在项目根目录 .env 更新登录令牌');
  }
  return /^bearer\s+/i.test(authorization) ? authorization : `bearer ${authorization}`;
}

function createClient() {
  const headers = {
    accept: 'application/json',
    'accept-language': 'zh-CN,zh;q=0.9',
    authorization: getAuthorization(),
    'content-type': 'application/json;charset=UTF-8',
    origin: 'https://connect.feibing.tech',
    referer: 'https://connect.feibing.tech/',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
  };
  if (process.env.VINCI_COOKIE) headers.cookie = process.env.VINCI_COOKIE;
  return axios.create({ headers, timeout: config.timeoutMs, validateStatus: () => true });
}

function assertApiSuccess(response, action) {
  if (response.status === 401 || response.status === 403) {
    throw new Error(`${action}失败：登录凭据无效或已过期（HTTP ${response.status}）`);
  }
  if (response.status < 200 || response.status >= 300) {
    throw new Error(`${action}失败：HTTP ${response.status}`);
  }
  if (response.data?.code !== 0) {
    throw new Error(`${action}失败：${response.data?.message || `code=${response.data?.code}`}`);
  }
  return response.data;
}

async function readWithRetry(action, requestFn, maxAttempts = 3) {
  let lastError;
  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      const response = await requestFn();
      if ((response.status === 429 || response.status >= 500) && attempt < maxAttempts) {
        throw new Error(`HTTP ${response.status}`);
      }
      return response;
    } catch (error) {
      lastError = error;
      if (attempt >= maxAttempts) break;
      const waitMs = attempt * 1000;
      console.warn(`${action}暂时失败（第 ${attempt}/${maxAttempts} 次）：${error.message}，${waitMs}ms 后重试`);
      await delay(waitMs);
    }
  }
  throw new Error(`${action}连续 ${maxAttempts} 次失败：${lastError?.message || '未知错误'}`);
}

function collectTargetCandidates(userData) {
  const candidates = [];
  const visit = (value, key = '') => {
    if (value == null) return;
    if (typeof value === 'string' || typeof value === 'number') {
      candidates.push({ key, value: String(value) });
      return;
    }
    if (Array.isArray(value)) {
      value.forEach((item) => visit(item, key));
      return;
    }
    if (typeof value === 'object') {
      Object.entries(value).forEach(([childKey, childValue]) => visit(childValue, childKey));
    }
  };
  visit(userData);
  return candidates;
}

async function resolveTargetCustomer(client) {
  const response = await readWithRetry(
    '查询目标客服',
    () => client.post(
      'https://vinci-api.feibing.tech/xc/v1/cps/token/contact/search',
      { queryWord: config.targetAccount, sellerId: config.sellerId },
    ),
  );
  const data = assertApiSuccess(response, '查询目标客服');
  const userData = data?.data?.user;
  const userIds = Array.isArray(userData?.userid) ? userData.userid.filter(Boolean).map(String) : [];
  if (userIds.length !== 1) {
    throw new Error(`账号 ${config.targetAccount} 必须唯一匹配 1 个客服，当前匹配 ${userIds.length} 个`);
  }
  const candidates = collectTargetCandidates(userData);
  const nameMatched = candidates.some(({ value }) => value.includes('青苹果'));
  const accountMatched = candidates.some(({ value }) => value === config.targetAccount);
  // contact/search 当前通常只返回 userid，不返回手机号和别名。
  // 因此硬校验采用“精确账号查询 + 唯一 userid”；若接口带回账号/别名，则额外记录交叉验证结果。
  return {
    userId: userIds[0],
    accountMatched,
    nameMatched,
    verification: nameMatched && accountMatched ? 'api-confirmed' : 'exact-account-unique-userid',
  };
}

async function getCurrentOwners(client, shopId) {
  const response = await readWithRetry(
    `查询门店 ${shopId} 当前客服`,
    () => client.get(
      `https://vinci-api.feibing.tech/sc/v1/sellers/${config.sellerId}/places`,
      {
        params: {
          current: 1,
          pageSize: 100,
          pageNum: 1,
          shopId,
          catalogId: '',
          types: 'DESK,SINGLE,CHANNEL,GROUP',
        },
      },
    ),
  );
  const data = assertApiSuccess(response, `查询门店 ${shopId} 当前客服`);
  const places = Array.isArray(data.data) ? data.data : [];
  const owners = new Set();
  for (const place of places) {
    try {
      const contact = place?.metadata?.contact ? JSON.parse(place.metadata.contact) : null;
      if (Array.isArray(contact?.owners)) {
        contact.owners.forEach((owner) => owner && owners.add(String(owner).trim()));
      }
    } catch (_) {
      // 单个桌台 contact 格式异常时跳过，最终由 places/owners 摘要暴露给预演检查。
    }
  }
  return { placeCount: places.length, owners: [...owners] };
}

async function updateContact(client, shopId, owners) {
  if (!Array.isArray(owners) || owners.length === 0) {
    throw new Error(`门店 ${shopId} 的目标客服为空，拒绝更新`);
  }
  const response = await client.post(
    `https://vinci-api.feibing.tech/sc/v1/sellers/${config.sellerId}/places/${shopId}/batch/update/contact`,
    {
      batchUpdateType: 0,
      shopIds: [shopId],
      customerService: JSON.stringify({ nickName: owners[0] }),
      owners,
    },
  );
  return assertApiSuccess(response, `更新门店 ${shopId} 客服`);
}

async function preflight(client, stores, target) {
  const records = [];
  for (let index = 0; index < stores.length; index += 1) {
    const store = stores[index];
    const current = await getCurrentOwners(client, store.shopId);
    if (current.placeCount === 0) {
      throw new Error(`门店 ${store.shopName} (${store.shopId}) 未查到桌台，终止整批任务`);
    }
    records.push({
      ...store,
      currentPlaceCount: current.placeCount,
      oldOwners: current.owners,
      targetUserId: target.userId,
    });
    console.log(`[${index + 1}/${stores.length}] ${store.shopName}: ${current.placeCount} 个桌台，当前客服 ${current.owners.length} 个`);
    await delay(config.requestDelayMs);
  }
  return records;
}

function writeJson(filePath, value) {
  ensureOutputDir();
  fs.writeFileSync(filePath, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

async function executeReplace(client, records) {
  const backupPath = path.join(config.outputDir, `备份_8月31日异常门店换客服_${timestamp()}.json`);
  writeJson(backupPath, {
    createdAt: new Date().toISOString(),
    sellerId: config.sellerId,
    targetAccount: config.targetAccount,
    targetName: config.targetName,
    records,
  });
  console.log(`写前备份已生成: ${backupPath}`);

  const results = [];
  for (let index = 0; index < records.length; index += 1) {
    const record = records[index];
    try {
      await updateContact(client, record.shopId, [record.targetUserId]);
      results.push({ shopId: record.shopId, shopName: record.shopName, status: 'accepted' });
      console.log(`[${index + 1}/${records.length}] 已受理: ${record.shopName}`);
    } catch (error) {
      results.push({ shopId: record.shopId, shopName: record.shopName, status: 'failed', reason: error.message });
      console.error(`[${index + 1}/${records.length}] 失败: ${record.shopName} - ${error.message}`);
    }
    await delay(config.requestDelayMs);
  }
  const resultPath = path.join(config.outputDir, `结果_8月31日异常门店换客服_${timestamp()}.json`);
  writeJson(resultPath, { backupPath, results });
  return { backupPath, resultPath, results };
}

async function rollback(client, backupPath) {
  if (!fs.existsSync(backupPath)) throw new Error(`备份文件不存在: ${backupPath}`);
  const backup = JSON.parse(fs.readFileSync(backupPath, 'utf8'));
  if (backup.sellerId !== config.sellerId || !Array.isArray(backup.records)) {
    throw new Error('备份文件格式或 sellerId 不匹配');
  }
  const results = [];
  for (let index = 0; index < backup.records.length; index += 1) {
    const record = backup.records[index];
    if (!Array.isArray(record.oldOwners) || record.oldOwners.length === 0) {
      results.push({ shopId: record.shopId, shopName: record.shopName, status: 'skipped', reason: '原客服为空' });
      continue;
    }
    try {
      await updateContact(client, record.shopId, record.oldOwners);
      results.push({ shopId: record.shopId, shopName: record.shopName, status: 'accepted' });
      console.log(`[${index + 1}/${backup.records.length}] 回滚已受理: ${record.shopName}`);
    } catch (error) {
      results.push({ shopId: record.shopId, shopName: record.shopName, status: 'failed', reason: error.message });
    }
    await delay(config.requestDelayMs);
  }
  const resultPath = path.join(config.outputDir, `结果_8月31日异常门店换客服_回滚_${timestamp()}.json`);
  writeJson(resultPath, { backupPath, results });
  return { resultPath, results };
}

function parseArgs(argv) {
  const rollbackArg = argv.find((arg) => arg.startsWith('--rollback='));
  return {
    checkInput: argv.includes('--check-input'),
    execute: argv.includes('--execute'),
    rollbackPath: rollbackArg ? rollbackArg.slice('--rollback='.length) : '',
  };
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const stores = readStores();
  console.log(`输入校验通过：${stores.length} 家门店，来源行为 5~21，目标客服 ${config.targetName} (${config.targetAccount})`);
  if (args.checkInput) return;

  const client = createClient();
  if (args.rollbackPath) {
    const rollbackPath = path.resolve(args.rollbackPath);
    const result = await rollback(client, rollbackPath);
    console.log(`回滚请求完成，结果: ${result.resultPath}`);
    return;
  }

  const target = await resolveTargetCustomer(client);
  console.log(`目标客服校验通过：精确账号 ${config.targetAccount} 唯一匹配 userID ${target.userId}`);
  if (target.verification !== 'api-confirmed') {
    console.log(`提示：搜索接口未返回客服别名，名称“${config.targetName}”按用户提供信息记录。`);
  }
  const records = await preflight(client, stores, target);
  const preflightPath = path.join(config.outputDir, `预演_8月31日异常门店换客服_${timestamp()}.json`);
  writeJson(preflightPath, {
    createdAt: new Date().toISOString(),
    dryRun: !args.execute,
    sellerId: config.sellerId,
    targetAccount: config.targetAccount,
    targetName: config.targetName,
    targetUserId: target.userId,
    records,
  });
  console.log(`预演清单已生成: ${preflightPath}`);

  if (!args.execute) {
    console.log('当前为预演模式，未调用任何写入接口。');
    return;
  }
  if (process.env.CUSTOMER_SERVICE_WRITE_CONFIRM !== '8月31日17家门店') {
    throw new Error('缺少二次写入确认，必须设置 CUSTOMER_SERVICE_WRITE_CONFIRM=8月31日17家门店');
  }
  const result = await executeReplace(client, records);
  const accepted = result.results.filter((item) => item.status === 'accepted').length;
  console.log(`正式更新请求完成：已受理 ${accepted}/${records.length}，结果: ${result.resultPath}`);
}

if (require.main === module) {
  main().catch((error) => {
    console.error(`执行失败: ${error.message}`);
    process.exit(1);
  });
}

module.exports = { config, readStores, validateStores, collectTargetCandidates, parseArgs, readWithRetry };
