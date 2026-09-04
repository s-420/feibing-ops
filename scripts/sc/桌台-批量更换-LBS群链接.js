/**
 * @对象    桌台
 * @动作    批量写
 * @风险    高
 * @来源    store-data-extractor/PiLiangGengHuanZuoTaiLBSQunLianJie.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

/**
 * 批量更换桌台码 / LBS 码的入群链接
 *
 * 输入 Excel：店铺ID / 入群链接
 * 处理范围：门店下名称包含「桌台」或「LBS」的桌台、渠道码。
 * 修改位置：metadata.contact.welMsg.attachments 中 type=link 的 link 字段。
 *
 * 使用前：
 *   1. 在 .env 配置 VINCI_AUTHORIZATION；
 *   2. 准备 input/门店ID和桌台LBS入群链接_茶瀑布.xlsx；
 *   3. 先保持 dryRun=true 预演，核对日志后再改成 false 真写。
 */

require('dotenv').config();

const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const axios = require('axios');

const config = {
  sellerId: 'wpV3cNDAAA3ZRQDQZt0UsR9wORkx0thg', // 茶瀑布
  excelPath: path.join(__dirname, '../input', '门店ID和桌台LBS入群链接_茶瀑布.xlsx'),
  shopIdCol: '店铺ID',
  linkCol: '入群链接',
  logPath: path.join(__dirname, '../output', '批量更换桌台LBS群链接_茶瀑布_执行日志.txt'),
  dryRun: true, // ★ 默认只预演；确认无误后手动改为 false
  mode: 'replace', // replace 替换 | rollback 回滚
  backupPath: '', // 回滚时填写备份文件路径
  requestDelayMs: 200,
  maxRetries: 3,
  retryDelayMs: 2000,
  concurrency: 2,
  headers: {
    accept: '*/*',
    'accept-language': 'zh-CN,zh;q=0.9',
    authorization: process.env.VINCI_AUTHORIZATION || '',
    origin: 'https://connect.feibing.tech',
    referer: 'https://connect.feibing.tech/',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/138.0.0.0 Safari/537.36',
  },
};

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function runWithConcurrency(items, concurrency, worker) {
  const results = new Array(items.length);
  let nextIndex = 0;
  const runWorker = async () => {
    while (nextIndex < items.length) {
      const i = nextIndex++;
      results[i] = await worker(items[i], i);
    }
  };
  const workerCount = Math.min(concurrency, items.length);
  await Promise.all(Array.from({ length: workerCount }, () => runWorker()));
  return results;
}

function logToFile(message) {
  try {
    fs.mkdirSync(path.dirname(config.logPath), { recursive: true });
    const ts = new Date().toLocaleString('zh-CN', {
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false,
    });
    fs.appendFileSync(config.logPath, `[${ts}] ${message}\n`, 'utf8');
  } catch (e) {
    console.error('写入日志失败:', e.message);
  }
}

function isTargetPlace(place) {
  return /桌台|lbs/i.test(String(place?.name || ''));
}

function parseContact(place) {
  const raw = place?.metadata?.contact;
  if (!raw) return null;
  return typeof raw === 'string' ? JSON.parse(raw) : raw;
}

function findLinkAttachment(contact) {
  return contact?.welMsg?.attachments?.find(item => item?.type === 'link') || null;
}

function getCurrentLink(place) {
  try {
    return findLinkAttachment(parseContact(place))?.link;
  } catch (e) {
    return undefined;
  }
}

function prepareLinkUpdate(placeDetail, newLink) {
  const contact = parseContact(placeDetail);
  if (!contact) throw new Error('无 metadata.contact');

  const linkAttachment = findLinkAttachment(contact);
  if (!linkAttachment) throw new Error('欢迎语中无 type=link 的附件');

  const oldLink = linkAttachment.link || '';
  linkAttachment.link = newLink;

  const parseMetadataJson = (key, fallback) => {
    const raw = placeDetail.metadata?.[key];
    if (!raw) return fallback;
    return typeof raw === 'string' ? JSON.parse(raw) : raw;
  };

  const tableAttachment = contact.welMsg?.attachments?.find(item => item?.type !== 'link');
  const body = {
    ...placeDetail,
    shopId: placeDetail.shop?.id || placeDetail.shopId,
    '$catalogs': placeDetail.catalog?.id || 'default',
    content: contact.welMsg?.content || '',
    table: tableAttachment,
    attachments: [linkAttachment],
    COVER: placeDetail.media,
    contactDTO: contact,
    poster: parseMetadataJson('poster', {}),
    tableStickers: parseMetadataJson('tableStickers', {}),
    posterId: 'default',
    tableStickersId: 'default',
    tagCheckBox: [],
  };

  return { body, oldLink };
}

async function getPlaces(shopId) {
  const r = await axios.get(`https://vinci-api.feibing.tech/sc/v1/sellers/${config.sellerId}/places`, {
    params: { current: 1, pageSize: 200, pageNum: 1, shopId, catalogId: '', types: 'DESK,SINGLE,CHANNEL,GROUP' },
    headers: config.headers,
    timeout: 20000,
  });
  if (r.data && r.data.code !== 0) {
    throw new Error(`查桌台失败(code=${r.data.code}): ${r.data.message || '未知错误'}（若为 401 请更新 .env token）`);
  }
  return r.data.data || [];
}

async function getPlaceDetail(placeId) {
  const r = await axios.get(`https://vinci-api.feibing.tech/sc/v1/sellers/${config.sellerId}/places/${placeId}`, {
    headers: config.headers,
    timeout: 60000,
  });
  if (r.data && r.data.code !== 0) {
    throw new Error(`查详情失败(code=${r.data.code}): ${r.data.message || '未知错误'}`);
  }
  return r.data.data;
}

async function putLink(placeId, newLink) {
  const detail = await getPlaceDetail(placeId);
  const { body, oldLink } = prepareLinkUpdate(detail, newLink);
  const r = await axios.put(`https://vinci-api.feibing.tech/sc/v1/sellers/${config.sellerId}/places/${placeId}`, body, {
    headers: { ...config.headers, 'content-type': 'application/json;charset=UTF-8' },
    timeout: 60000,
  });
  return { result: r.data, oldLink };
}

function makeBackupPath() {
  const ts = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14);
  return path.join(__dirname, '../output', `备份_桌台LBS群链接_${ts}.jsonl`);
}

function backupRecord(record) {
  fs.appendFileSync(config.backupPath, JSON.stringify(record) + '\n', 'utf8');
}

async function processPlace(place, newLink) {
  const placeName = place.name || '(未命名)';
  const placeId = place.id;
  const currentLink = getCurrentLink(place);

  if (currentLink === undefined) {
    logToFile(`「${placeName}」(${placeId}) - 无可替换的群链接，跳过`);
    return 'skipped';
  }
  if (currentLink === newLink) {
    logToFile(`「${placeName}」(${placeId}) - 已是目标链接，跳过`);
    return 'skipped';
  }

  if (config.dryRun) {
    logToFile(`「${placeName}」(${placeId}) - 【预演】${currentLink || '(空)'} → ${newLink}`);
    return 'dryrun';
  }

  backupRecord({ placeId, placeName, oldLink: currentLink, newLink });
  for (let attempt = 1; attempt <= config.maxRetries; attempt++) {
    try {
      const { result } = await putLink(placeId, newLink);
      if (result && result.code === 0) {
        logToFile(`「${placeName}」(${placeId}) - 修改成功${attempt > 1 ? `（第${attempt}次尝试）` : ''}`);
        return 'success';
      }
      logToFile(`「${placeName}」(${placeId}) - 修改失败: ${result?.message || '未知原因'}`);
      return 'failed';
    } catch (e) {
      if (attempt === config.maxRetries) {
        logToFile(`「${placeName}」(${placeId}) - 重试${config.maxRetries}次仍失败: ${e.message}`);
        return 'failed';
      }
      await delay(config.retryDelayMs * attempt);
    }
  }
  return 'failed';
}

async function rollback() {
  if (!config.backupPath || !fs.existsSync(config.backupPath)) {
    throw new Error(`找不到备份文件: ${config.backupPath || '(未填写 backupPath)'}`);
  }
  const records = fs.readFileSync(config.backupPath, 'utf8')
    .split('\n').filter(Boolean).map(line => JSON.parse(line));
  let success = 0;
  let failed = 0;
  for (const record of records) {
    try {
      const { result } = await putLink(record.placeId, record.oldLink);
      if (result && result.code === 0) success++;
      else failed++;
    } catch (e) {
      failed++;
      logToFile(`「${record.placeName}」(${record.placeId}) - 回滚失败: ${e.message}`);
    }
    if (config.requestDelayMs > 0) await delay(config.requestDelayMs);
  }
  logToFile(`回滚结束：成功 ${success} / 失败 ${failed}`);
  console.log(`回滚结束：成功 ${success} / 失败 ${failed}`);
}

function readInputRows() {
  if (!fs.existsSync(config.excelPath)) throw new Error(`输入文件不存在: ${config.excelPath}`);
  const workbook = xlsx.readFile(config.excelPath);
  const rows = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
  const mapping = new Map();
  for (const row of rows) {
    const shopId = String(row[config.shopIdCol] ?? '').trim();
    const link = String(row[config.linkCol] ?? '').trim();
    if (!shopId || !link) continue;
    if (mapping.has(shopId) && mapping.get(shopId) !== link) {
      throw new Error(`店铺ID ${shopId} 存在多个不同的入群链接`);
    }
    mapping.set(shopId, link);
  }
  return mapping;
}

async function main() {
  try {
    if (config.mode === 'rollback') {
      await rollback();
      return;
    }
    if (!config.headers.authorization) throw new Error('缺少 VINCI_AUTHORIZATION，请先配置 .env');

    const shopLinks = readInputRows();
    if (shopLinks.size === 0) throw new Error('输入表中没有有效的「店铺ID / 入群链接」数据');
    if (!config.dryRun) config.backupPath = makeBackupPath();

    const tasks = [];
    for (const [shopId, newLink] of shopLinks) {
      const places = await getPlaces(shopId);
      const targets = places.filter(isTargetPlace);
      logToFile(`门店 ${shopId}：共 ${places.length} 条记录，命中桌台/LBS ${targets.length} 条`);
      targets.forEach(place => tasks.push({ place, newLink }));
    }

    const summary = { success: 0, failed: 0, skipped: 0, dryrun: 0 };
    let done = 0;
    await runWithConcurrency(tasks, config.concurrency, async ({ place, newLink }) => {
      const status = await processPlace(place, newLink);
      summary[status]++;
      done++;
      process.stdout.write(`\r进度 ${done}/${tasks.length}：成功 ${summary.success} / 预演 ${summary.dryrun} / 跳过 ${summary.skipped} / 失败 ${summary.failed}`);
    });

    console.log('\n执行完成');
    logToFile(`执行完成：成功 ${summary.success} / 预演 ${summary.dryrun} / 跳过 ${summary.skipped} / 失败 ${summary.failed}`);
    if (!config.dryRun) console.log(`备份: ${config.backupPath}`);
    console.log(`日志: ${config.logPath}`);
  } catch (e) {
    logToFile(`程序执行出错: ${e.message}`);
    console.error(`程序执行出错: ${e.message}`);
    process.exitCode = 1;
  }
}

if (require.main === module) main();

module.exports = { isTargetPlace, getCurrentLink, prepareLinkUpdate, findLinkAttachment };
