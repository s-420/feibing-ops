/**
 * @对象    门店
 * @动作    导出
 * @风险    低
 * @来源    store-data-extractor/DaoChuWeiXinMenDianHaoRuQunLianJieChaPuBu.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

// 用来导出门店身上的微信门店号的入群链接，只限于茶瀑布的企业。
const axios = require('axios');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const CONFIG = {
  inputPath: path.join(__dirname, '../input/茶瀑布专用导出入群链接表格.xlsx'),
  outputPath: path.join(__dirname, '../output/茶瀑布微信门店号入群链接.xlsx'),
  inputStoreIdColumn: '店铺id',
  inputStoreNameColumn: '店铺名称',
  keySuffix: '_pre_h5_config_wx',
  baseUrl: 'https://vinci-api.feibing.tech/sc/v1/sellers/wpV3cNDAAA3ZRQDQZt0UsR9wORkx0thg/configs',
  concurrency: 10,
  maxRetries: 3,
  retryDelayMs: 1000,
  requestDelayMs: 200,
  requestTimeoutMs: 20000,
  headers: {
    accept: '*/*',
    'accept-language': 'zh-CN,zh;q=0.9',
    authorization: 'bearer __VINCI_TOKEN__',
    origin: 'https://connect.feibing.tech',
    referer: 'https://connect.feibing.tech/',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36'
  }
};

const GROUP_LINK_PATTERN = /^https:\/\/work\.weixin\.qq\.com\/gm\/[^"'\\\s<>()]+$/;
const GROUP_LINK_SEARCH_PATTERN = /https:\/\/work\.weixin\.qq\.com\/gm\/[^"'\\\s<>()]+/;

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function normalizeStoreId(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function storeIdFromKey(key, fallbackStoreId) {
  const normalizedKey = normalizeStoreId(key);
  if (normalizedKey.endsWith(CONFIG.keySuffix)) {
    return normalizedKey.slice(0, -CONFIG.keySuffix.length);
  }
  return fallbackStoreId;
}

function readStores() {
  if (!fs.existsSync(CONFIG.inputPath)) {
    throw new Error('找不到输入文件：' + CONFIG.inputPath);
  }

  const workbook = XLSX.readFile(CONFIG.inputPath);
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error('输入表格没有工作表');
  }

  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    header: 1,
    defval: ''
  });

  if (rows.length === 0) {
    throw new Error('输入表格为空');
  }

  const headers = rows[0].map(value => normalizeStoreId(value));
  const storeIdColumnIndex = headers.indexOf(CONFIG.inputStoreIdColumn);
  const storeNameColumnIndex = headers.indexOf(CONFIG.inputStoreNameColumn);
  if (storeIdColumnIndex === -1) {
    throw new Error('输入表格缺少“' + CONFIG.inputStoreIdColumn + '”列');
  }
  if (storeNameColumnIndex === -1) {
    throw new Error('输入表格缺少“' + CONFIG.inputStoreNameColumn + '”列');
  }

  const stores = rows
    .slice(1)
    .map(row => ({
      storeId: normalizeStoreId(row[storeIdColumnIndex]),
      storeName: normalizeStoreId(row[storeNameColumnIndex])
    }))
    .filter(store => store.storeId);

  if (stores.length === 0) {
    throw new Error('“' + CONFIG.inputStoreIdColumn + '”列没有有效门店');
  }

  return stores;
}

function readStoreIds() {
  return readStores().map(store => store.storeId);
}

function findGroupLinkInNode(node) {
  if (Array.isArray(node)) {
    for (const item of node) {
      const link = findGroupLinkInNode(item);
      if (link) return link;
    }
    return '';
  }

  if (!node || typeof node !== 'object') {
    return '';
  }

  if (typeof node.path === 'string' && GROUP_LINK_PATTERN.test(node.path)) {
    return node.path;
  }

  for (const value of Object.values(node)) {
    const link = findGroupLinkInNode(value);
    if (link) return link;
  }

  return '';
}

function extractGroupLink(value) {
  if (typeof value !== 'string' || value.trim() === '') {
    return '';
  }

  try {
    const parsedValue = JSON.parse(value);
    const structuredLink = findGroupLinkInNode(parsedValue);
    if (structuredLink) return structuredLink;
  } catch {
    // 兼容历史脏数据：JSON解析失败时继续使用正则提取。
  }

  return value.match(GROUP_LINK_SEARCH_PATTERN)?.[0] || '';
}

function isRetryableError(error) {
  if (!error.response) return true;
  const status = error.response.status;
  return status === 408 || status === 429 || status >= 500;
}

async function requestStoreConfig(storeId) {
  const key = storeId + CONFIG.keySuffix;
  let lastError;

  for (let attempt = 1; attempt <= CONFIG.maxRetries; attempt += 1) {
    try {
      const response = await axios.get(CONFIG.baseUrl, {
        params: {
          type: 'MINI_APP_STYLE_TYPE',
          key
        },
        headers: CONFIG.headers,
        timeout: CONFIG.requestTimeoutMs
      });

      const body = response.data;
      if (!body || body.code !== 0) {
        throw new Error('接口返回失败：' + (body?.message || '未知错误') + '（code: ' + body?.code + '）');
      }

      if (!Array.isArray(body.data) || body.data.length === 0) {
        return null;
      }

      return body.data.find(item => item?.key === key) || body.data[0];
    } catch (error) {
      lastError = error;
      if (attempt >= CONFIG.maxRetries || !isRetryableError(error)) {
        break;
      }

      await delay(CONFIG.retryDelayMs);
    }
  }

  throw lastError;
}

async function processStore(store) {
  const { storeId, storeName } = store;
  try {
    const storeConfig = await requestStoreConfig(storeId);
    if (!storeConfig) {
      return {
        storeId,
        storeName,
        groupLink: '',
        status: 'no_config'
      };
    }

    const groupLink = extractGroupLink(storeConfig.value);
    return {
      storeId: storeIdFromKey(storeConfig.key, storeId),
      storeName,
      groupLink,
      status: groupLink ? 'success' : 'no_link'
    };
  } catch (error) {
    return {
      storeId,
      storeName,
      groupLink: '',
      status: 'failed',
      error: error.response?.data?.message || error.message
    };
  }
}

// 固定并发池：任一请求完成后立即补充下一条，并保持输出顺序与输入表格一致。
async function runWithConcurrency(items, concurrency, worker, onCompleted) {
  const results = new Array(items.length);
  let nextIndex = 0;

  const runWorker = async () => {
    while (nextIndex < items.length) {
      const currentIndex = nextIndex;
      nextIndex += 1;

      results[currentIndex] = await worker(items[currentIndex], currentIndex);
      onCompleted?.(results[currentIndex], currentIndex);

      if (nextIndex < items.length && CONFIG.requestDelayMs > 0) {
        await delay(CONFIG.requestDelayMs);
      }
    }
  };

  const workerCount = Math.min(concurrency, items.length);
  await Promise.all(
    Array.from({ length: workerCount }, () => runWorker())
  );

  return results;
}

async function exportResults(results) {
  fs.mkdirSync(path.dirname(CONFIG.outputPath), { recursive: true });

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'DaoChuWeiXinMenDianHaoRuQunLianJie';
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet('入群链接', {
    views: [{ state: 'frozen', ySplit: 1 }]
  });

  worksheet.columns = [
    { header: '门店ID', key: 'storeId', width: 30 },
    { header: '店铺名称', key: 'storeName', width: 42 },
    { header: '入群链接', key: 'groupLink', width: 72 }
  ];

  for (const result of results) {
    worksheet.addRow({
      storeId: result.storeId,
      storeName: result.storeName,
      groupLink: result.groupLink
    });
  }

  const headerRow = worksheet.getRow(1);
  headerRow.height = 24;
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF1F4E78' }
  };

  worksheet.autoFilter = {
    from: 'A1',
    to: 'C' + Math.max(1, results.length + 1)
  };

  worksheet.getColumn('storeId').numFmt = '@';
  worksheet.getColumn('storeName').numFmt = '@';
  worksheet.getColumn('groupLink').numFmt = '@';
  worksheet.getColumn('storeId').alignment = { vertical: 'middle' };
  worksheet.getColumn('storeName').alignment = { vertical: 'middle' };
  worksheet.getColumn('groupLink').alignment = { vertical: 'middle' };

  await workbook.xlsx.writeFile(CONFIG.outputPath);
}

function summarize(results) {
  return results.reduce(
    (summary, result) => {
      summary[result.status] = (summary[result.status] || 0) + 1;
      return summary;
    },
    { success: 0, no_link: 0, no_config: 0, failed: 0 }
  );
}

async function main() {
  console.log('=====================================================');
  console.log('开始导出茶瀑布微信门店号入群链接');
  console.log('输入文件：' + CONFIG.inputPath);

  const stores = readStores();
  console.log('门店总数：' + stores.length);
  console.log('并发数：' + CONFIG.concurrency + '（完成一条立即补充下一条）');

  let completedCount = 0;
  const results = await runWithConcurrency(
    stores,
    CONFIG.concurrency,
    processStore,
    result => {
      completedCount += 1;
      const suffix = result.status === 'failed' ? '，错误：' + result.error : '';
      console.log(
        '[' + completedCount + '/' + stores.length + '] ' +
        result.storeId + '：' + result.status + suffix
      );
    }
  );

  await exportResults(results);

  const summary = summarize(results);
  console.log('=====================================================');
  console.log('导出完成：' + CONFIG.outputPath);
  console.log(
    '成功提取：' + summary.success +
    ' | 无入群链接：' + summary.no_link +
    ' | 无配置：' + summary.no_config +
    ' | 请求失败：' + summary.failed
  );
}

if (require.main === module) {
  main().catch(error => {
    console.error('脚本执行失败：', error.message);
    process.exitCode = 1;
  });
}

module.exports = {
  CONFIG,
  extractGroupLink,
  findGroupLinkInNode,
  readStoreIds,
  readStores,
  runWithConcurrency,
  storeIdFromKey
};

