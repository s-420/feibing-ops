/**
 * @对象    桌台
 * @动作    批量写
 * @风险    高
 * @来源    store-data-extractor/PiLiangTiHuanZuoTaiHuanYingYuNeiRong.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

/**
 * 批量替换桌台欢迎语文字（不修改发送/不发送开关）。
 *
 * 默认只预演：
 *   npm run PiLiangTiHuanZuoTaiHuanYingYuNeiRong
 * 真正执行：
 *   npm run PiLiangTiHuanZuoTaiHuanYingYuNeiRong -- --execute
 * 指定欢迎语文件：
 *   npm run PiLiangTiHuanZuoTaiHuanYingYuNeiRong -- --content-file input/桌台欢迎语.txt
 * 只处理一个门店：
 *   npm run PiLiangTiHuanZuoTaiHuanYingYuNeiRong -- --shop-id 门店ID
 * 回滚：
 *   npm run PiLiangTiHuanZuoTaiHuanYingYuNeiRong -- --rollback output/备份_桌台欢迎语内容_时间.jsonl
 */

require('dotenv').config();

const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const axios = require('axios');

function getArg(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : '';
}

const rollbackPath = getArg('--rollback');
const config = {
  sellerId: getArg('--seller-id') || 'wpV3cNDAAA3ZRQDQZt0UsR9wORkx0thg',
  excelPath: path.resolve(getArg('--excel') || path.join(__dirname, '../input', '门店ID和客服手机号_茶瀑布.xlsx')),
  contentPath: path.resolve(getArg('--content-file') || path.join(__dirname, '../input', '桌台欢迎语.txt')),
  shopIdCol: '店铺ID',
  shopId: getArg('--shop-id'),
  logPath: path.join(__dirname, '../output', '批量替换桌台欢迎语内容_茶瀑布_执行日志.txt'),
  dryRun: !process.argv.includes('--execute'),
  mode: rollbackPath ? 'rollback' : 'replace',
  backupPath: rollbackPath ? path.resolve(rollbackPath) : '',
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
      const index = nextIndex++;
      results[index] = await worker(items[index], index);
    }
  };
  await Promise.all(Array.from({ length: Math.min(concurrency, items.length) }, () => runWorker()));
  return results;
}

function logToFile(message) {
  fs.mkdirSync(path.dirname(config.logPath), { recursive: true });
  const timestamp = new Date().toLocaleString('zh-CN', { hour12: false });
  fs.appendFileSync(config.logPath, `[${timestamp}] ${message}\n`, 'utf8');
}

function assertAuthorization() {
  if (!config.headers.authorization.trim()) {
    throw new Error('缺少 VINCI_AUTHORIZATION，请先在 .env 中配置登录令牌');
  }
}

function readNewContent() {
  if (!fs.existsSync(config.contentPath)) {
    throw new Error(`找不到欢迎语文件: ${config.contentPath}`);
  }
  // 文本文件通常以换行结尾，而后台保存时会自动去掉这个末尾换行。
  // 读取时同步归一化，避免复查误判并重复写入。
  const content = fs.readFileSync(config.contentPath, 'utf8')
    .replace(/^\uFEFF/, '')
    .replace(/\r?\n$/, '');
  if (!content.trim()) {
    throw new Error(`欢迎语文件内容为空: ${config.contentPath}`);
  }
  return content;
}

async function getPlaces(shopId) {
  const response = await axios.get(`https://vinci-api.feibing.tech/sc/v1/sellers/${config.sellerId}/places`, {
    params: { current: 1, pageSize: 200, pageNum: 1, shopId, catalogId: '', types: 'DESK,SINGLE,CHANNEL,GROUP' },
    headers: config.headers,
    timeout: 20000,
  });
  if (response.data && response.data.code !== 0) {
    throw new Error(`查桌台失败(code=${response.data.code}): ${response.data.message || '未知错误'}`);
  }
  return response.data.data || [];
}

async function getPlaceDetail(placeId) {
  const response = await axios.get(`https://vinci-api.feibing.tech/sc/v1/sellers/${config.sellerId}/places/${placeId}`, {
    headers: config.headers,
    timeout: 60000,
  });
  if (response.data && response.data.code !== 0) {
    throw new Error(`查详情失败(code=${response.data.code}): ${response.data.message || '未知错误'}`);
  }
  return response.data.data;
}

function getCurrentContent(place) {
  try {
    return JSON.parse(place.metadata?.contact).welMsg?.content;
  } catch (error) {
    return undefined;
  }
}

async function putContent(placeId, targetContent) {
  const detail = await getPlaceDetail(placeId);
  const contact = JSON.parse(detail.metadata.contact);
  const tableStickers = JSON.parse(detail.metadata.tableStickers);
  const poster = JSON.parse(detail.metadata.poster);

  contact.welMsg.content = targetContent;

  const body = {
    ...detail,
    shopId: detail.shop.id,
    '$catalogs': detail.catalog?.id || 'default',
    content: targetContent,
    table: contact.welMsg.attachments[0],
    attachments: [contact.welMsg.attachments[1]],
    COVER: detail.media,
    contactDTO: contact,
    poster,
    tableStickers,
    posterId: 'default',
    tableStickersId: 'default',
    tagCheckBox: [],
  };

  const response = await axios.put(`https://vinci-api.feibing.tech/sc/v1/sellers/${config.sellerId}/places/${placeId}`, body, {
    headers: { ...config.headers, 'content-type': 'application/json;charset=UTF-8' },
    timeout: 60000,
  });
  return response.data;
}

function makeBackupPath() {
  const timestamp = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14);
  return path.join(__dirname, '../output', `备份_桌台欢迎语内容_${timestamp}.jsonl`);
}

async function putWithRetry(placeId, content) {
  for (let attempt = 1; attempt <= config.maxRetries; attempt++) {
    try {
      const result = await putContent(placeId, content);
      if (result && result.code === 0) return result;
      throw new Error(result?.message || '接口返回未知错误');
    } catch (error) {
      if (attempt === config.maxRetries) throw error;
      await delay(config.retryDelayMs * attempt);
    }
  }
}

async function processPlace(place, targetContent) {
  const name = place.name || '(未命名)';
  const currentContent = getCurrentContent(place);
  if (currentContent === undefined) {
    logToFile(`桌台「${name}」(${place.id}) - 无 welMsg/content，跳过`);
    return 'skipped';
  }
  if (currentContent === targetContent) {
    logToFile(`桌台「${name}」(${place.id}) - 欢迎语已是目标内容，跳过`);
    return 'skipped';
  }
  if (config.dryRun) {
    logToFile(`桌台「${name}」(${place.id}) - 【预演】欢迎语将被替换（未写入）`);
    return 'dryrun';
  }

  fs.appendFileSync(config.backupPath, JSON.stringify({
    placeId: place.id,
    placeName: name,
    oldContent: currentContent,
    newContent: targetContent,
  }) + '\n', 'utf8');

  await putWithRetry(place.id, targetContent);
  logToFile(`桌台「${name}」(${place.id}) - ✅ 欢迎语替换成功`);
  return 'success';
}

async function rollback() {
  if (!config.backupPath || !fs.existsSync(config.backupPath)) {
    throw new Error(`找不到备份文件: ${config.backupPath || '(未指定)'}`);
  }
  const records = fs.readFileSync(config.backupPath, 'utf8')
    .split('\n')
    .filter(line => line.trim())
    .map(line => JSON.parse(line));
  let success = 0;
  let failed = 0;
  for (const record of records) {
    try {
      await putWithRetry(record.placeId, record.oldContent);
      success++;
      logToFile(`桌台「${record.placeName}」(${record.placeId}) - 回滚成功`);
    } catch (error) {
      failed++;
      logToFile(`桌台「${record.placeName}」(${record.placeId}) - 回滚失败: ${error.message}`);
    }
  }
  console.log(`回滚结束: 成功 ${success} / 失败 ${failed}`);
}

function printProgress(done, total, summary) {
  const percentage = total ? ((done / total) * 100).toFixed(1) : '100.0';
  process.stdout.write(`\r进度: ${done}/${total} (${percentage}%) [成功 ${summary.success} / 预演 ${summary.dryrun} / 跳过 ${summary.skipped} / 失败 ${summary.failed}]`);
}

async function main() {
  try {
    assertAuthorization();
    if (config.mode === 'rollback') {
      await rollback();
      return;
    }

    const targetContent = readNewContent();
    if (!config.dryRun) config.backupPath = makeBackupPath();

    let shopIds;
    if (config.shopId) {
      shopIds = [config.shopId.trim()];
    } else {
      if (!fs.existsSync(config.excelPath)) throw new Error(`找不到门店表格: ${config.excelPath}`);
      const workbook = xlsx.readFile(config.excelPath);
      const rows = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
      shopIds = [...new Set(rows.map(row => String(row[config.shopIdCol] ?? '').trim()).filter(Boolean))];
      if (!shopIds.length) throw new Error(`门店表格中没有有效的「${config.shopIdCol}」`);
    }

    console.log(`模式: ${config.dryRun ? '预演（不会写入）' : '正式执行'}`);
    console.log(`门店数: ${shopIds.length}，欢迎语文件: ${config.contentPath}`);
    logToFile(`===== 开始：${config.dryRun ? '预演' : '正式执行'}，门店 ${shopIds.length} 个 =====`);

    const allPlaces = [];
    for (const shopId of shopIds) {
      try {
        allPlaces.push(...await getPlaces(shopId));
      } catch (error) {
        logToFile(`门店 ${shopId} - 查桌台失败: ${error.message}`);
      }
    }

    const summary = { success: 0, dryrun: 0, skipped: 0, failed: 0 };
    let done = 0;
    await runWithConcurrency(allPlaces, config.concurrency, async place => {
      try {
        const result = await processPlace(place, targetContent);
        summary[result]++;
      } catch (error) {
        summary.failed++;
        logToFile(`桌台「${place.name || '(未命名)'}」(${place.id}) - ❌ 替换失败: ${error.message}`);
      }
      printProgress(++done, allPlaces.length, summary);
    });

    console.log('\n');
    console.log(`汇总: 成功 ${summary.success} / 预演 ${summary.dryrun} / 跳过 ${summary.skipped} / 失败 ${summary.failed}`);
    if (config.backupPath) console.log(`备份: ${config.backupPath}`);
    console.log(`日志: ${config.logPath}`);
  } catch (error) {
    console.error(`程序执行出错: ${error.message}`);
    process.exitCode = 1;
  }
}

main();
