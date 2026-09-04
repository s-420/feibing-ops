/**
 * @对象    门店
 * @动作    批量写
 * @风险    高
 * @来源    store-data-extractor/TiHuanMenDianRuQunLianJie.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

/**
 * 替换门店「入群链接」（安全版）
 *
 * 与旧脚本 XiuGaiWeiXinMenDianHaoLink.js 的核心区别：
 *   旧脚本：只有当门店当前的旧链接「严格等于」硬编码的 config.targetLink 时才替换，
 *           否则静默跳过 —— 因此要求所有门店的旧链接都得是同一个模板链接，有隐患。
 *   本脚本：不再依赖 targetLink。只认两件事：
 *           ① 固定位置 value.json.list[3].children[0].path；
 *           ② 该位置的值必须「长得像企业微信群链接」（https://work.weixin.qq.com/gm/...）。
 *           只要格式匹配就替换成 Excel 里该门店的新链接；HTML 里所有群链接一并替换。
 *           配置里除链接以外的其他内容一律不动（不做整份覆盖）。
 *
 * 兜底机制：
 *   ① 写前备份：真正写入（dryRun=false）时，每更新一个门店之前，先把它的
 *      「完整原配置 value」追加写进 output/备份_替换入群链接_时间戳.jsonl。
 *      备份失败会中止该门店的更新（保证「没有备份就绝不改动」）。
 *   ② 回滚：把 config.mode 改成 'rollback'、config.backupPath 填上备份文件路径，
 *      脚本会读备份文件，把每个门店的原配置逐个 PUT 回去，一键恢复。
 *
 * 用法：
 *   1. 准备 input/门店ID和入群链接.xlsx（列：门店名称 / 店铺ID / 入群链接）
 *   2. 确认下方 ACTIVE_BRAND（'qingxiang' 轻享 或 'chapubu' 茶瀑布）
 *   3. 先保持 dryRun = true 跑一遍，看 output 日志里「预演」结果是否符合预期
 *   4. 确认无误后把 dryRun 改成 false，再跑一遍真正写入（会自动生成备份文件）
 *   5. 需要回滚时：把 mode 改成 'rollback'、backupPath 指向第4步生成的备份文件，再跑一遍
 */

require('dotenv').config();

const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const axios = require('axios');

// 群链接格式校验（与导出脚本 DaoChuWeiXinMenDianHaoRuQunLianJieChaPuBu.js 保持一致）
const GROUP_LINK_PATTERN = /^https:\/\/work\.weixin\.qq\.com\/gm\/[^"'\\\s<>()]+$/; // 单值校验
const GROUP_LINK_GLOBAL = /https:\/\/work\.weixin\.qq\.com\/gm\/[^"'\\\s<>()]+/g;   // HTML 全局替换

// ======================= 品牌配置（二选一）=======================
// authorization 优先读取 .env 里的 VINCI_AUTHORIZATION（项目统一凭据，有效期约24小时），
// 这里保留硬编码值仅作兜底；正常情况下请更新根目录 .env 文件，而不是改这里。
const BRANDS = {
  // 轻享
  qingxiang: {
    baseUrl: 'https://vinci-api.feibing.tech/sc/v1/sellers/wwa9c5a585540b115b/configs',
    excelPath: path.join(__dirname, '../input', '门店ID和入群链接.xlsx'),
    logPath: path.join(__dirname, '../output', '替换入群链接_执行日志.txt'),
    authorization: 'bearer __VINCI_TOKEN__',
  },
  // 茶瀑布
  chapubu: {
    baseUrl: 'https://vinci-api.feibing.tech/sc/v1/sellers/wpV3cNDAAA3ZRQDQZt0UsR9wORkx0thg/configs',
    excelPath: path.join(__dirname, '../input', '门店ID和入群链接_茶瀑布.xlsx'),
    logPath: path.join(__dirname, '../output', '替换入群链接_茶瀑布_执行日志.txt'),
    authorization: 'bearer __VINCI_TOKEN__',
  },
};

// ★★ 跑哪个品牌就改成哪个：'qingxiang' 或 'chapubu'
const ACTIVE_BRAND = 'chapubu';

const brand = BRANDS[ACTIVE_BRAND];

const config = {
  baseUrl: brand.baseUrl,
  excelPath: brand.excelPath,
  logPath: brand.logPath,
  keySuffix: '_pre_h5_config_wx',
  // ★★ 预演开关：true = 只记录「会怎么改」，不真正调用 PUT 写入。
  //     先跑一遍 true 看日志，确认无误后再改成 false 真正写入。
  dryRun: false,
  // 运行模式：'replace' 替换（默认） | 'rollback' 回滚
  mode: 'replace',
  // 回滚模式时填：要恢复的备份文件路径（即替换模式跑完后生成的 备份_*.jsonl）
  backupPath: '',
  requestDelayMs: 200, // 每次请求间隔，避免限流
  headers: {
    'accept': '*/*',
    'accept-language': 'zh-CN,zh;q=0.9',
    'authorization': process.env.VINCI_AUTHORIZATION || brand.authorization,
    'origin': 'https://connect.feibing.tech',
    'referer': 'https://connect.feibing.tech/',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36',
  },
};

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// 读取 Excel 文件
function readExcelFile(filePath) {
  try {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    return xlsx.utils.sheet_to_json(worksheet);
  } catch (error) {
    const errorMsg = `读取Excel文件失败: ${error.message}`;
    logToFile(errorMsg);
    console.error(errorMsg);
    throw error;
  }
}

// 记录日志到 TXT 文件
function logToFile(message) {
  try {
    if (!fs.existsSync(path.dirname(config.logPath))) {
      fs.mkdirSync(path.dirname(config.logPath), { recursive: true });
    }

    const timestamp = new Date().toLocaleString('zh-CN', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false
    });

    const logEntry = `[${timestamp}] ${message}\n`;
    fs.appendFileSync(config.logPath, logEntry, 'utf8');
  } catch (error) {
    console.error('写入日志失败:', error.message);
  }
}

// 生成带时间戳的备份文件路径（JSONL 格式：一行一个门店）
function makeBackupPath() {
  const ts = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14); // YYYYMMDDHHmmss
  return path.join(__dirname, '../output', `备份_替换入群链接_${ts}.jsonl`);
}

// 兜底①：写前备份门店的「完整原配置」。备份失败则抛错，中止本次更新（保证不丢原数据）。
function backupOriginal(configDetails, storeId, newLink) {
  const record = {
    storeId,
    configId: configDetails.id,
    type: configDetails.type,
    key: configDetails.key,
    value: configDetails.value, // 原始完整 value（JSON 字符串）
    newLink,                    // 记录改成了什么，便于事后核对
    backupTime: new Date().toISOString()
  };
  fs.appendFileSync(config.backupPath, JSON.stringify(record) + '\n', 'utf8');
  logToFile(`门店 ${storeId} - 已备份原配置（${record.configId}）`);
}

// 获取配置详情
async function getConfigDetails(storeId) {
  try {
    const url = `${config.baseUrl}?type=MINI_APP_STYLE_TYPE&key=${storeId}${config.keySuffix}`;
    logToFile(`正在获取门店 ${storeId} 的配置详情...`);

    const response = await axios.get(url, { headers: config.headers });

    if (response.data.code !== 0) {
      const errorMsg = `获取配置失败: ${response.data.message}`;
      logToFile(errorMsg);
      throw new Error(errorMsg);
    }

    if (!response.data.data || response.data.data.length === 0) {
      const errorMsg = '未找到配置数据';
      logToFile(errorMsg);
      throw new Error(errorMsg);
    }

    return response.data.data[0];
  } catch (error) {
    const errorMsg = `获取门店 ${storeId} 配置失败: ${error.response?.data?.message || error.message}`;
    logToFile(errorMsg);
    throw error;
  }
}

// 更新配置
async function updateConfig(configId, updatedConfig) {
  try {
    const url = `${config.baseUrl}/${configId}`;
    const putHeaders = {
      ...config.headers,
      'content-type': 'application/json;charset=UTF-8'
    };

    logToFile(`正在更新配置 ${configId}...`);
    const response = await axios.put(url, updatedConfig, { headers: putHeaders });

    return response.data;
  } catch (error) {
    const errorMsg = `更新配置 ${configId} 失败: ${error.response?.data?.message || error.message}`;
    logToFile(errorMsg);
    throw error;
  }
}

// 替换链接并准备更新数据 —— 只认「位置 + 群链接格式」，不认 targetLink
function prepareUpdateData(originalConfig, newLink, storeId) {
  const valueObj = JSON.parse(originalConfig.value);
  let replaceCount = 0;
  const originalLinks = []; // 存储替换前的原链接，便于核对/回滚

  // 1. JSON 固定位置 value.json.list[3].children[0].path：
  //    只要当前位置的值长得像群链接就替换；否则告警并跳过（不盲目覆盖非链接内容）
  if (valueObj.json && valueObj.json.list && valueObj.json.list.length >= 4) {
    const targetItem = valueObj.json.list[3];
    if (targetItem.children && targetItem.children[0]) {
      const oldPath = targetItem.children[0].path;
      if (typeof oldPath === 'string' && GROUP_LINK_PATTERN.test(oldPath)) {
        originalLinks.push(oldPath);
        logToFile(`门店 ${storeId} - JSON位置替换: ${oldPath} -> ${newLink}`);
        targetItem.children[0].path = newLink;
        replaceCount++;
      } else {
        logToFile(`门店 ${storeId} - ⚠ JSON位置的值不是群链接，跳过（原值: ${oldPath}）`);
      }
    } else {
      logToFile(`门店 ${storeId} - ⚠ JSON位置缺少 children[0]，跳过JSON部分`);
    }
  } else {
    logToFile(`门店 ${storeId} - ⚠ 配置结构不符合预期（list 不足4项），跳过JSON部分`);
  }

  // 2. HTML：所有群链接一并替换
  if (typeof valueObj.html === 'string' && valueObj.html !== '') {
    const matches = valueObj.html.match(GROUP_LINK_GLOBAL);
    if (matches && matches.length > 0) {
      originalLinks.push(...matches);
      logToFile(`门店 ${storeId} - HTML替换，共${matches.length}处群链接 -> ${newLink}`);
      valueObj.html = valueObj.html.replace(GROUP_LINK_GLOBAL, newLink);
      replaceCount += matches.length;
    } else {
      logToFile(`门店 ${storeId} - HTML 中未找到群链接`);
    }
  }

  if (replaceCount === 0) {
    logToFile(`门店 ${storeId} - 未找到任何群链接，不执行更新`);
    return { updateData: null, originalLinks: [] };
  }

  return {
    updateData: {
      type: originalConfig.type,
      key: originalConfig.key,
      value: JSON.stringify(valueObj)
    },
    originalLinks
  };
}

// 处理单个门店
async function processStore(store) {
  const storeId = store['店铺ID'];
  const newLink = store['入群链接'];

  logToFile(`\n开始处理门店: ${storeId}`);
  console.log(`处理中: ${storeId}`);

  // 获取配置详情
  const configDetails = await getConfigDetails(storeId);

  // 准备更新数据
  const { updateData, originalLinks } = prepareUpdateData(configDetails, newLink, storeId);

  // 无需更新
  if (!updateData) {
    logToFile(`门店 ${storeId} - 处理完成，未进行更新`);
    console.log(`已完成: ${storeId} (未更新)`);
    return 'skipped';
  }

  // 预演模式：只记录，不写入
  if (config.dryRun) {
    logToFile(`门店 ${storeId} - 【预演】原链接: ${originalLinks.join('; ')}，将替换为: ${newLink}（未写入）`);
    console.log(`已完成: ${storeId} (预演，未写入)`);
    return 'dryrun';
  }

  // 兜底①：写前备份完整原配置（备份失败会抛错，从而中止本次更新）
  backupOriginal(configDetails, storeId, newLink);

  // 执行更新
  const updateResult = await updateConfig(configDetails.id, updateData);

  if (updateResult.code === 0) {
    logToFile(`门店 ${storeId} - 更新成功`);
    logToFile(`门店 ${storeId} - 原始链接: ${originalLinks.join('; ')}`);
    logToFile(`门店 ${storeId} - 替换为: ${newLink}`);
    console.log(`已完成: ${storeId} (成功)`);
    return 'success';
  }

  logToFile(`门店 ${storeId} - 更新失败: ${updateResult.message || '更新失败，未知原因'}`);
  console.log(`已完成: ${storeId} (失败)`);
  return 'failed';
}

// 兜底②：回滚 —— 读备份文件，把每个门店的原配置逐个 PUT 回去
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
      console.log(`跳过无法解析的行`);
      fail++;
      continue;
    }

    try {
      const body = { type: record.type, key: record.key, value: record.value };
      const result = await updateConfig(record.configId, body);
      if (result.code === 0) {
        ok++;
        logToFile(`门店 ${record.storeId} - 回滚成功`);
        console.log(`回滚成功: ${record.storeId}`);
      } else {
        fail++;
        logToFile(`门店 ${record.storeId} - 回滚失败: ${result.message || '未知原因'}`);
        console.log(`回滚失败: ${record.storeId}`);
      }
    } catch (e) {
      fail++;
      logToFile(`门店 ${record.storeId} - 回滚失败: ${e.message}`);
      console.log(`回滚失败: ${record.storeId}`);
    }

    if (config.requestDelayMs > 0) {
      await delay(config.requestDelayMs);
    }
  }

  logToFile(`===== 回滚结束: 成功 ${ok} / 失败 ${fail} =====`);
  console.log(`回滚结束: 成功 ${ok} / 失败 ${fail}`);
}

// 主函数
async function main() {
  const summary = { success: 0, failed: 0, skipped: 0, dryrun: 0, error: 0 };

  try {
    // 回滚模式：走回滚流程，不执行替换
    if (config.mode === 'rollback') {
      await rollback();
      return;
    }

    logToFile('===== 程序开始执行 =====');
    console.log(`程序开始执行...（品牌: ${ACTIVE_BRAND}，模式: ${config.mode}，预演: ${config.dryRun ? '开' : '关'}）`);

    // 替换模式 + 真正写入时，先创建备份文件路径
    if (!config.dryRun) {
      config.backupPath = makeBackupPath();
      logToFile(`本次备份文件: ${config.backupPath}`);
    }

    // 检查 input 目录是否存在
    if (!fs.existsSync(path.dirname(config.excelPath))) {
      fs.mkdirSync(path.dirname(config.excelPath), { recursive: true });
      const errorMsg = '已创建input目录，请将Excel文件放入该目录后重新运行';
      logToFile(errorMsg);
      console.error(errorMsg);
      process.exit(1);
    }

    // 读取 Excel 数据
    logToFile('正在读取Excel文件...');
    const stores = readExcelFile(config.excelPath);
    logToFile(`成功读取 ${stores.length} 条门店数据`);
    console.log(`共读取到 ${stores.length} 条门店数据，开始处理...`);

    // 逐个处理门店
    for (const store of stores) {
      if (!store['店铺ID'] || !store['入群链接']) {
        const msg = `跳过缺少信息的行 - 店铺ID: ${store['店铺ID'] || '未知'}, 入群链接: ${store['入群链接'] || '未知'}`;
        logToFile(msg);
        console.log(`跳过: ${store['店铺ID'] || '未知'} (信息不完整)`);
        summary.skipped++;
        continue;
      }

      try {
        const result = await processStore(store);
        summary[result] = (summary[result] || 0) + 1;
      } catch (error) {
        logToFile(`门店 ${store['店铺ID']} - 处理失败: ${error.message}`);
        console.log(`已完成: ${store['店铺ID']} (失败)`);
        summary.error++;
      }

      if (config.requestDelayMs > 0) {
        await delay(config.requestDelayMs);
      }
    }

    logToFile('===== 程序执行结束 =====');
    logToFile(`汇总: 成功 ${summary.success} / 失败 ${summary.failed} / 未更新 ${summary.skipped} / 预演 ${summary.dryrun} / 异常 ${summary.error}`);
    logToFile('');
    console.log(`\n程序执行完成，详细日志已保存到: ${config.logPath}`);
    console.log(`汇总: 成功 ${summary.success} / 失败 ${summary.failed} / 未更新 ${summary.skipped} / 预演 ${summary.dryrun} / 异常 ${summary.error}`);
  } catch (error) {
    const errorMsg = `程序执行出错: ${error.message}`;
    logToFile(errorMsg);
    console.error(errorMsg);
    process.exit(1);
  }
}

// 启动程序
main();
