/**
 * @对象    桌台
 * @动作    批量写
 * @风险    高
 * @来源    store-data-extractor/PiLiangXiuGaiZuoTaiHuanYingYu.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

/**
 * 批量修改桌台欢迎语（不发送 → 发送）
 *
 * 业务：客户要求把一批门店的桌台欢迎语，从「不发送欢迎语」批量改成「发送欢迎语」。
 *
 * 关键字段：noticeType（点单卡片 platformWeapp）
 *   - 0 = 发送欢迎语
 *   - 1 = 不发送欢迎语
 *
 * 位置（PUT body 中，需同步改「点单卡片」的两处）：
 *   - contactDTO.welMsg.attachments[0].noticeType（点单卡片 platformWeapp）
 *   - table.noticeType（同一点单卡片的表单字段）
 *   （入群链接 link 的 noticeType 保持 0，不动）
 *
 * 流程（每个桌台）：
 *   ① GET  places/{id}        拿完整桌台数据
 *   ② 解析 metadata.contact，改点单卡片的 noticeType → 目标值
 *   ③ PUT  places/{id}        提交（body = 完整数据 + 表单字段 + 修改后的 noticeType）
 *
 * 输入：门店列表 Excel（含「店铺ID」列）。对每个门店下所有桌台处理。
 */

require('dotenv').config();

const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const axios = require('axios');

const config = {
  sellerId: 'wpV3cNDAAA3ZRQDQZt0UsR9wORkx0thg', // 茶瀑布
  // 输入：门店列表（含「店铺ID」列）
  excelPath: path.join(__dirname, '../input', '门店ID和客服手机号_茶瀑布.xlsx'),
  shopIdCol: '店铺ID',
  logPath: path.join(__dirname, '../output', '批量修改桌台欢迎语_茶瀑布_执行日志.txt'),
  targetNoticeType: 0, // 0=发送欢迎语，1=不发送欢迎语（点单卡片 platformWeapp 的 noticeType）
  dryRun: false, // ★ 预演开关
  mode: 'replace', // 运行模式：replace 替换 | rollback 回滚
  backupPath: '', // 回滚时填：要恢复的备份文件路径
  requestDelayMs: 200,
  maxRetries: 3, // 失败自动重试次数（网络超时/锁竞争等可重试错误）
  retryDelayMs: 2000, // 重试退避基础延迟（毫秒），第 n 次重试等待 n * retryDelayMs
  concurrency: 2, // 并发数（温和，避免数据库锁竞争）
  headers: {
    'accept': '*/*',
    'accept-language': 'zh-CN,zh;q=0.9',
    'authorization': process.env.VINCI_AUTHORIZATION || '',
    'origin': 'https://connect.feibing.tech',
    'referer': 'https://connect.feibing.tech/',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/138.0.0.0 Safari/537.36',
  },
};

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// 简单并发池（固定并发数，任一完成即补充下一条）
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
    if (!fs.existsSync(path.dirname(config.logPath))) {
      fs.mkdirSync(path.dirname(config.logPath), { recursive: true });
    }
    const ts = new Date().toLocaleString('zh-CN', {
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false
    });
    fs.appendFileSync(config.logPath, `[${ts}] ${message}\n`, 'utf8');
  } catch (e) {
    console.error('写入日志失败:', e.message);
  }
}

// 查门店下所有桌台
async function getPlaces(shopId) {
  const r = await axios.get(`https://vinci-api.feibing.tech/sc/v1/sellers/${config.sellerId}/places`, {
    params: { current: 1, pageSize: 200, pageNum: 1, shopId, catalogId: '', types: 'DESK,SINGLE,CHANNEL,GROUP' },
    headers: config.headers, timeout: 20000,
  });
  if (r.data && r.data.code !== 0) {
    throw new Error(`查桌台失败(code=${r.data.code}): ${r.data.message || '未知错误'}（若为 401 请更新 .env 的 token）`);
  }
  return r.data.data || [];
}

// 查单个桌台详情
async function getPlaceDetail(placeId) {
  const r = await axios.get(`https://vinci-api.feibing.tech/sc/v1/sellers/${config.sellerId}/places/${placeId}`, {
    headers: config.headers, timeout: 60000,
  });
  if (r.data && r.data.code !== 0) {
    throw new Error(`查详情失败(code=${r.data.code}): ${r.data.message || '未知错误'}（若为 401 请更新 .env 的 token）`);
  }
  return r.data.data;
}

// 取当前点单卡片的 noticeType
function getCurrentNoticeType(d) {
  try {
    const c = JSON.parse(d.metadata?.contact);
    return c.welMsg?.attachments?.[0]?.noticeType;
  } catch (e) {
    return undefined;
  }
}

// PUT 修改 noticeType
async function putNoticeType(placeId, targetNoticeType) {
  const d = await getPlaceDetail(placeId);
  const contact = JSON.parse(d.metadata.contact);
  const tableStickers = JSON.parse(d.metadata.tableStickers);
  const poster = JSON.parse(d.metadata.poster);

  // 改点单卡片 noticeType
  contact.welMsg.attachments[0].noticeType = targetNoticeType;

  const body = {
    ...d,
    shopId: d.shop.id,
    '$catalogs': d.catalog?.id || 'default',
    content: contact.welMsg.content,
    table: contact.welMsg.attachments[0],
    attachments: [contact.welMsg.attachments[1]],
    COVER: d.media,
    contactDTO: contact,
    poster,
    tableStickers,
    posterId: 'default',
    tableStickersId: 'default',
    tagCheckBox: [],
  };
  if (body.table) body.table.noticeType = targetNoticeType;

  const r = await axios.put(`https://vinci-api.feibing.tech/sc/v1/sellers/${config.sellerId}/places/${placeId}`, body, {
    headers: { ...config.headers, 'content-type': 'application/json;charset=UTF-8' },
    timeout: 60000,
  });
  return r.data;
}

// 生成备份文件路径
function makeBackupPath() {
  const ts = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14);
  return path.join(__dirname, '../output', `备份_桌台欢迎语_${ts}.jsonl`);
}

// 备份：记录修改前的 noticeType（回滚依据）
function backupRecord(record) {
  fs.appendFileSync(config.backupPath, JSON.stringify(record) + '\n', 'utf8');
}

// 处理单个桌台
async function processPlace(place) {
  const name = place.name || '(未命名)';
  const placeId = place.id;

  // 列表接口已含完整 metadata.contact，直接读 noticeType，不必再 GET 详情
  const current = getCurrentNoticeType(place);

  if (current === undefined) {
    logToFile(`桌台「${name}」(${placeId}) - 无 welMsg/欢迎语，跳过`);
    console.log(`跳过: ${name} (无欢迎语)`);
    return 'skipped';
  }

  if (current === config.targetNoticeType) {
    logToFile(`桌台「${name}」(${placeId}) - 已是目标状态(noticeType=${current})，无需修改`);
    console.log(`跳过: ${name} (已是目标状态)`);
    return 'skipped';
  }

  logToFile(`桌台「${name}」(${placeId}) - 当前 noticeType=${current} → 目标 ${config.targetNoticeType}`);

  if (config.dryRun) {
    logToFile(`桌台「${name}」- 【预演】将 noticeType ${current} → ${config.targetNoticeType}（未写入）`);
    console.log(`预演: ${name} (${current}→${config.targetNoticeType})`);
    return 'dryrun';
  }

  // 真写：备份原值一次，然后带重试地 PUT
  backupRecord({ placeId, placeName: name, oldNoticeType: current, newNoticeType: config.targetNoticeType });

  for (let attempt = 1; attempt <= config.maxRetries; attempt++) {
    try {
      const result = await putNoticeType(placeId, config.targetNoticeType);
      if (result && result.code === 0) {
        logToFile(`桌台「${name}」- ✅ 修改成功: noticeType ${current} → ${config.targetNoticeType}${attempt > 1 ? `（第${attempt}次尝试）` : ''}`);
        console.log(`成功: ${name} (${current}→${config.targetNoticeType})`);
        return 'success';
      }
      // 业务失败（code !== 0）
      logToFile(`桌台「${name}」- ❌ 修改失败: ${result?.message || '未知原因'}`);
      console.log(`失败: ${name}`);
      return 'failed';
    } catch (e) {
      // 网络超时 / 锁竞争 / TLS 断连 等可重试错误
      if (attempt < config.maxRetries) {
        const wait = config.retryDelayMs * attempt;
        logToFile(`桌台「${name}」- 第${attempt}次失败(${e.message})，${wait}ms 后重试...`);
        await delay(wait);
      } else {
        logToFile(`桌台「${name}」- ❌ 重试${config.maxRetries}次仍失败: ${e.message}`);
        console.log(`失败: ${name} (重试${config.maxRetries}次)`);
        return 'failed';
      }
    }
  }
  return 'failed';
}

// 回滚：读备份文件，把每个桌台的 noticeType 改回原值
async function rollback() {
  if (!config.backupPath || !fs.existsSync(config.backupPath)) {
    const msg = `找不到备份文件: ${config.backupPath || '(未填写 backupPath)'}`;
    logToFile(msg);
    console.error(msg);
    process.exit(1);
  }

  const lines = fs.readFileSync(config.backupPath, 'utf8').split('\n').filter(l => l.trim());
  logToFile(`===== 回滚开始，共 ${lines.length} 条记录 =====`);
  console.log(`回滚开始，共 ${lines.length} 条记录...`);

  let ok = 0, fail = 0;
  for (const line of lines) {
    let record;
    try { record = JSON.parse(line); } catch (e) { logToFile(`跳过无法解析的行`); fail++; continue; }
    try {
      const result = await putNoticeType(record.placeId, record.oldNoticeType);
      if (result && result.code === 0) {
        ok++;
        logToFile(`桌台「${record.placeName}」- 回滚成功: noticeType → ${record.oldNoticeType}`);
        console.log(`回滚成功: ${record.placeName}`);
      } else {
        fail++;
        logToFile(`桌台「${record.placeName}」- 回滚失败: ${result?.message || '未知原因'}`);
        console.log(`回滚失败: ${record.placeName}`);
      }
    } catch (e) {
      fail++;
      logToFile(`桌台「${record.placeName}」- 回滚失败: ${e.message}`);
      console.log(`回滚失败: ${record.placeName}`);
    }
    if (config.requestDelayMs > 0) await delay(config.requestDelayMs);
  }

  logToFile(`===== 回滚结束: 成功 ${ok} / 失败 ${fail} =====`);
  console.log(`回滚结束: 成功 ${ok} / 失败 ${fail}`);
}

// 覆盖式进度条
function printProgress(done, total, summary) {
  const pct = total > 0 ? ((done / total) * 100).toFixed(1) : '100.0';
  process.stdout.write(`\r进度: ${done}/${total} (${pct}%)  [成功 ${summary.success} / 跳过 ${summary.skipped} / 失败 ${summary.failed}]`);
}

async function main() {
  const summary = { success: 0, failed: 0, skipped: 0, dryrun: 0, error: 0 };

  try {
    // 回滚模式
    if (config.mode === 'rollback') {
      await rollback();
      return;
    }

    logToFile('===== 程序开始执行 =====');
    console.log(`程序开始执行...（茶瀑布，预演: ${config.dryRun ? '开' : '关'}，目标 noticeType: ${config.targetNoticeType}）`);

    // 真写时初始化备份文件
    if (!config.dryRun) {
      config.backupPath = makeBackupPath();
      logToFile(`本次备份文件: ${config.backupPath}`);
    }

    const wb = xlsx.readFile(config.excelPath);
    const rows = xlsx.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
    // 去重店铺ID
    const shopIds = [...new Set(rows.map(r => String(r[config.shopIdCol] ?? '').trim()).filter(Boolean))];
    logToFile(`读取到 ${shopIds.length} 个门店`);

    // 1. 预查：收集所有桌台 + 统计需要修改的数量
    console.log(`共 ${shopIds.length} 个门店，预查桌台列表...`);
    const allPlaces = [];
    for (const shopId of shopIds) {
      try {
        const places = await getPlaces(shopId);
        for (const place of places) allPlaces.push(place);
      } catch (e) {
        logToFile(`门店 ${shopId} - 查桌台失败: ${e.message}`);
      }
    }
    let needChange = 0;
    for (const place of allPlaces) {
      const cur = getCurrentNoticeType(place);
      if (cur !== undefined && cur !== config.targetNoticeType) needChange++;
    }
    logToFile(`共 ${allPlaces.length} 个桌台，其中 ${needChange} 个需要修改`);
    console.log(`共 ${allPlaces.length} 个桌台，其中 ${needChange} 个需要修改，开始处理...`);

    // 2. 并发处理 + 实时进度
    let done = 0;
      await runWithConcurrency(allPlaces, config.concurrency, async (place) => {
      try {
        const r = await processPlace(place);
        summary[r] = (summary[r] || 0) + 1;
      } catch (e) {
        logToFile(`桌台「${place.name}」- 处理失败: ${e.message}`);
        summary.error++;
      }
      done++;
      printProgress(done, allPlaces.length, summary);
    });

    console.log('\n');
    logToFile('===== 程序执行结束 =====');
    logToFile(`汇总: 成功 ${summary.success} / 失败 ${summary.failed} / 跳过 ${summary.skipped} / 预演 ${summary.dryrun} / 异常 ${summary.error}`);
    console.log(`汇总: 成功 ${summary.success} / 失败 ${summary.failed} / 跳过 ${summary.skipped} / 预演 ${summary.dryrun} / 异常 ${summary.error}`);
    console.log(`日志: ${config.logPath}`);
  } catch (e) {
    const msg = `程序执行出错: ${e.message}`;
    logToFile(msg);
    console.error(msg);
    process.exit(1);
  }
}

main();
