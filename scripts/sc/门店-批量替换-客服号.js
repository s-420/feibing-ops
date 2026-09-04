/**
 * @对象    门店
 * @动作    批量写
 * @风险    高
 * @来源    store-data-extractor/TiHuanMenDianKeFuHao.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

/**
 * 替换门店绑定的客服号（安全版）
 *
 * 与入群链接替换脚本 TiHuanMenDianRuQunLianJie.js 同一套设计：
 *   ① 预演 dryRun（默认 true，只看不改）
 *   ② 写前备份（真写前查出门店当前客服存进备份文件，备份失败中止该门店更新）
 *   ③ 一键回滚（mode='rollback' 读备份文件恢复原客服）
 *
 * 品牌：茶瀑布（sellerId = wpV3cNDAAA3ZRQDQZt0UsR9wORkx0thg）
 * 输入：input/门店ID和客服手机号_茶瀑布.xlsx（列：门店名称 / 店铺ID / 客服手机号）
 *
 * 接口（复用 PiLiangXiuGaiZuoTaiDuiYinKeFu.js 的逻辑）：
 *   查当前客服：GET  /places?shopId={店铺ID}  → place.metadata.contact.owners[]
 *   改客服：    POST /places/{店铺ID}/batch/update/contact
 *               { batchUpdateType:0, shopIds:[店铺ID], customerService:{"nickName":"手机号"}, owners:["手机号"] }
 */

require('dotenv').config();

const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const axios = require('axios');

const config = {
  baseUrl: 'https://vinci-api.feibing.tech/sc/v1/sellers/wpV3cNDAAA3ZRQDQZt0UsR9wORkx0thg', // 茶瀑布
  excelPath: path.join(__dirname, '../input', '门店ID和客服手机号_茶瀑布.xlsx'),
  logPath: path.join(__dirname, '../output', '替换客服号_茶瀑布_执行日志.txt'),
  // ★★ 预演开关：true = 只记录「会怎么改」，不真正调用 POST 写入
  dryRun: false,
  // 运行模式：'replace' 替换（默认） | 'rollback' 回滚
  mode: 'replace',
  // 回滚模式时填：要恢复的备份文件路径
  backupPath: '',
  requestDelayMs: 300, // 每次请求间隔，避免限流
  headers: {
    'accept': '*/*',
    'accept-language': 'zh-CN,zh;q=0.9',
    'authorization': process.env.VINCI_AUTHORIZATION || '',
    'origin': 'https://connect.feibing.tech',
    'referer': 'https://connect.feibing.tech/',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36',
  },
};

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function logToFile(message) {
  try {
    if (!fs.existsSync(path.dirname(config.logPath))) {
      fs.mkdirSync(path.dirname(config.logPath), { recursive: true });
    }
    const timestamp = new Date().toLocaleString('zh-CN', {
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false
    });
    fs.appendFileSync(config.logPath, `[${timestamp}] ${message}\n`, 'utf8');
  } catch (error) {
    console.error('写入日志失败:', error.message);
  }
}

function readExcel(filePath) {
  const workbook = xlsx.readFile(filePath);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  return xlsx.utils.sheet_to_json(worksheet);
}

function makeBackupPath() {
  const ts = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14);
  return path.join(__dirname, '../output', `备份_替换客服号_${ts}.jsonl`);
}

// 查门店当前客服（返回去重后的 owners 列表，即当前绑定的客服号集合）
async function getCurrentOwners(shopId) {
  const url = `${config.baseUrl}/places`;
  const response = await axios.get(url, {
    params: {
      current: 1,
      pageSize: 100,
      pageNum: 1,
      shopId,
      catalogId: '',
      types: 'DESK,SINGLE,CHANNEL,GROUP',
    },
    headers: config.headers,
    timeout: 20000,
  });

  if (response.data.code !== 0) {
    throw new Error(`查询门店客服失败: ${response.data.message}`);
  }

  const items = response.data.data || [];
  const ownersSet = new Set();
  for (const item of items) {
    try {
      if (item.metadata?.contact) {
        const contact = JSON.parse(item.metadata.contact);
        if (Array.isArray(contact.owners)) {
          contact.owners.forEach(o => { if (o) ownersSet.add(String(o).trim()); });
        }
      }
    } catch (e) {
      // 单个 place 的 contact 解析失败，跳过
    }
  }
  return {
    places: items.length,
    owners: [...ownersSet],
  };
}

// 手机号 → 客服 userID（茶瀑布系统用 userID 标识客服，直接传手机号不生效）
async function phoneToUserid(phone) {
  const response = await axios.post(
    'https://vinci-api.feibing.tech/xc/v1/cps/token/contact/search',
    { queryWord: phone, sellerId: 'wpV3cNDAAA3ZRQDQZt0UsR9wORkx0thg' },
    {
      headers: { ...config.headers, 'content-type': 'application/json;charset=UTF-8' },
      timeout: 20000,
    }
  );
  if (response.data.code !== 0) {
    throw new Error(`手机号 ${phone} 查询客服失败: ${response.data.message}`);
  }
  const userid = response.data?.data?.user?.userid?.[0];
  if (!userid) {
    throw new Error(`手机号 ${phone} 未找到对应客服`);
  }
  return userid;
}

// 改门店客服（复用 PiLiangXiuGaiZuoTaiDuiYinKeFu.js 的请求体）
async function updateContact(shopId, customerServiceValue) {
  const requestData = {
    batchUpdateType: 0,
    shopIds: [shopId],
    customerService: JSON.stringify({ nickName: customerServiceValue }),
    owners: [customerServiceValue],
  };

  const response = await axios.post(
    `${config.baseUrl}/places/${shopId}/batch/update/contact`,
    requestData,
    {
      headers: { ...config.headers, 'content-type': 'application/json;charset=UTF-8' },
      timeout: 30000,
      validateStatus: () => true,
    }
  );

  return response.data;
}

// 处理单个门店（替换模式）
async function processStore(store) {
  const name = store['门店名称'];
  const shopId = store['店铺ID'];
  const phone = store['客服手机号'];

  logToFile(`\n开始处理门店: ${name} (${shopId})`);
  console.log(`处理中: ${name}`);

  // 查当前客服
  const { places, owners } = await getCurrentOwners(shopId);
  const oldOwnersStr = owners.length ? owners.join('; ') : '(无)';
  logToFile(`门店 ${name} - 当前客服号(${places}个桌台): ${oldOwnersStr}`);
  logToFile(`门店 ${name} - 目标客服号: ${phone}`);

  // 预演模式：只记录，不写入
  if (config.dryRun) {
    logToFile(`门店 ${name} - 【预演】将把客服号从 [${oldOwnersStr}] 替换为 ${phone}（未写入）`);
    console.log(`已完成: ${name} (预演，未写入)`);
    return 'dryrun';
  }

  // 兜底：写前备份（备份失败会抛错，中止本次更新）
  backupOriginal({ name, shopId, oldOwners: owners, newPhone: phone });

  // 执行更新（先手机号→userID，再传 userID 替换）
  const userid = await phoneToUserid(phone);
  logToFile(`门店 ${name} - 手机号 ${phone} 对应客服 userID: ${userid}`);

  const result = await updateContact(shopId, userid);
  if (result && result.code === 0 && result.message === '成功') {
    logToFile(`门店 ${name} - 更新成功: [${oldOwnersStr}] -> ${phone}`);
    console.log(`已完成: ${name} (成功)`);
    return 'success';
  }
  const msg = result?.message?.toString().slice(0, 80) || '未知原因';
  logToFile(`门店 ${name} - 更新失败(code=${result?.code}): ${msg}`);
  console.log(`已完成: ${name} (失败)`);
  return 'failed';
}

// 兜底：写前备份
function backupOriginal(record) {
  fs.appendFileSync(config.backupPath, JSON.stringify(record) + '\n', 'utf8');
  logToFile(`门店 ${record.name} - 已备份原客服号（${record.oldOwners.length} 个）`);
}

// 兜底：回滚 —— 读备份文件，把每个门店的原客服号逐个恢复
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
    try {
      record = JSON.parse(line);
    } catch (e) {
      logToFile(`跳过无法解析的行: ${line.slice(0, 80)}`);
      fail++;
      continue;
    }

    try {
      // 用备份的原客服号（owners[0]）恢复
      if (record.oldOwners && record.oldOwners.length > 1) {
        logToFile(`门店 ${record.name} - ⚠ 原绑定多个客服(${record.oldOwners.length}个)，回滚仅恢复到第一个: ${record.oldOwners[0]}`);
      }
      const restoreValue = record.oldOwners && record.oldOwners[0];
      if (!restoreValue) {
        logToFile(`门店 ${record.name} - 无备份客服号，跳过`);
        fail++;
        continue;
      }
      const result = await updateContact(record.shopId, restoreValue);
      if (result && result.code === 0 && result.message === '成功') {
        ok++;
        logToFile(`门店 ${record.name} - 回滚成功（恢复为 ${restoreValue}）`);
        console.log(`回滚成功: ${record.name}`);
      } else {
        fail++;
        logToFile(`门店 ${record.name} - 回滚失败: ${result?.message || '未知原因'}`);
        console.log(`回滚失败: ${record.name}`);
      }
    } catch (e) {
      fail++;
      logToFile(`门店 ${record.name} - 回滚失败: ${e.message}`);
      console.log(`回滚失败: ${record.name}`);
    }

    if (config.requestDelayMs > 0) {
      await delay(config.requestDelayMs);
    }
  }

  logToFile(`===== 回滚结束: 成功 ${ok} / 失败 ${fail} =====`);
  console.log(`回滚结束: 成功 ${ok} / 失败 ${fail}`);
}

async function main() {
  const summary = { success: 0, failed: 0, dryrun: 0, error: 0 };

  try {
    // 回滚模式
    if (config.mode === 'rollback') {
      await rollback();
      return;
    }

    logToFile('===== 程序开始执行 =====');
    console.log(`程序开始执行...（品牌: 茶瀑布，模式: ${config.mode}，预演: ${config.dryRun ? '开' : '关'}）`);

    if (!config.dryRun) {
      config.backupPath = makeBackupPath();
      logToFile(`本次备份文件: ${config.backupPath}`);
    }

    const stores = readExcel(config.excelPath);
    logToFile(`成功读取 ${stores.length} 条门店数据`);
    console.log(`共读取到 ${stores.length} 条门店数据，开始处理...`);

    for (const store of stores) {
      if (!store['店铺ID'] || !store['客服手机号']) {
        logToFile(`跳过缺少信息的行 - 门店: ${store['门店名称'] || '未知'}`);
        continue;
      }

      try {
        const result = await processStore(store);
        summary[result] = (summary[result] || 0) + 1;
      } catch (error) {
        logToFile(`门店 ${store['门店名称']} - 处理失败: ${error.message}`);
        console.log(`已完成: ${store['门店名称']} (失败)`);
        summary.error++;
      }

      if (config.requestDelayMs > 0) {
        await delay(config.requestDelayMs);
      }
    }

    logToFile('===== 程序执行结束 =====');
    logToFile(`汇总: 成功 ${summary.success} / 失败 ${summary.failed} / 预演 ${summary.dryrun} / 异常 ${summary.error}`);
    console.log(`\n程序执行完成，详细日志已保存到: ${config.logPath}`);
    console.log(`汇总: 成功 ${summary.success} / 失败 ${summary.failed} / 预演 ${summary.dryrun} / 异常 ${summary.error}`);
  } catch (error) {
    const errorMsg = `程序执行出错: ${error.message}`;
    logToFile(errorMsg);
    console.error(errorMsg);
    process.exit(1);
  }
}

main();
