/**
 * @对象    渠道欢迎语
 * @动作    批量写
 * @风险    高
 * @输入    已人工复核并完成原始数据备份的 .xlsx 文件目录
 * @输出    output/渠道欢迎语批量导入_时间.json
 * @验收    每批 /files 与 /batch/update/place/welmsg 均返回 code=0
 * @来源    store-data-extractor/PiLiangDaoRuQuDaoHuanYingYuExcel.js
 * @脱敏    sellerId、Authorization、Cookie 均由参数或环境变量注入
 */

/**
 * 上传并批量导入渠道欢迎语 Excel。
 *
 * 默认 dry-run：
 *   node scripts/sc/渠道欢迎语-批量导入-Excel.js \
 *     --input-dir=input/待导入 --place-name=渠道名称 --place-type=DESK
 * 正式执行（必须确认原始数据已备份）：
 *   node scripts/sc/渠道欢迎语-批量导入-Excel.js \
 *     --input-dir=input/待导入 --place-name=渠道名称 --place-type=DESK \
 *     --backup-confirmed --execute
 * 断点续传：追加 --start=2；限制批数：追加 --limit=1。
 * 已上传文件直接提交：追加 --excel-url=https://.../file.xlsx。
 *
 * 相同文件重复导入的边界：若服务端执行的是按同一对象覆盖、文件内容与目标
 * 完全一致，最终业务数据通常不变；但仍会新增文件/任务记录，也可能重复触发
 * 服务端副作用。因此脚本不会自动把“相同文件”当作无条件安全的幂等操作，
 * 应优先使用 --start 或 --excel-url 避免重复上传、重复提交。
 */

require('dotenv').config();

const fs = require('fs');
const path = require('path');
const axios = require('axios');

function getArg(name) {
  const args = process.argv.slice(2);
  const index = args.indexOf(name);
  if (index >= 0) return args[index + 1] || '';
  const prefixed = args.find(arg => arg.startsWith(`${name}=`));
  return prefixed ? prefixed.slice(name.length + 1) : '';
}

function parsePositiveInteger(name, fallback, allowZero = false) {
  const raw = getArg(name);
  const value = raw === '' ? fallback : Number(raw);
  const minimum = allowZero ? 0 : 1;
  if (!Number.isInteger(value) || value < minimum) throw new Error(`${name} 参数不合法`);
  return value;
}

function parseOptions() {
  const sellerId = getArg('--seller-id') || process.env.SELLER_ID || '';
  const inputDir = getArg('--input-dir');
  const excelUrl = getArg('--excel-url');
  const placeName = getArg('--place-name');
  const placeType = getArg('--place-type');
  if (!sellerId) throw new Error('缺少 --seller-id 或 SELLER_ID');
  if (!excelUrl && !inputDir) throw new Error('缺少 --input-dir（或使用 --excel-url）');
  if (!placeName) throw new Error('缺少 --place-name');
  if (!placeType) throw new Error('缺少 --place-type；请以浏览器实际请求为准，例如 DESK');
  return {
    sellerId,
    inputDir: inputDir ? path.resolve(inputDir) : '',
    excelUrl,
    placeName,
    placeType,
    start: parsePositiveInteger('--start', 1),
    limit: parsePositiveInteger('--limit', 0, true),
    execute: process.argv.includes('--execute'),
    uploadOnly: process.argv.includes('--upload-only'),
    backupConfirmed: process.argv.includes('--backup-confirmed'),
  };
}

function getHeaders(contentType) {
  const authorization = String(process.env.VINCI_AUTHORIZATION || '').trim();
  if (!authorization) throw new Error('缺少 VINCI_AUTHORIZATION');
  const headers = {
    accept: '*/*',
    authorization: /^bearer\s+/i.test(authorization) ? authorization : `bearer ${authorization}`,
    origin: 'https://connect.feibing.tech',
    referer: 'https://connect.feibing.tech/',
  };
  const cookie = String(process.env.VINCI_COOKIE || '').trim();
  if (cookie) headers.cookie = cookie;
  if (contentType) headers['content-type'] = contentType;
  return headers;
}

function listExcelFiles(inputDir, start = 1, limit = 0) {
  if (!fs.existsSync(inputDir)) throw new Error(`找不到目录：${inputDir}`);
  const files = fs.readdirSync(inputDir)
    .filter(name => name.toLowerCase().endsWith('.xlsx'))
    .sort((a, b) => a.localeCompare(b, 'zh-CN', { numeric: true }))
    .map(name => path.join(inputDir, name))
    .slice(start - 1);
  const selected = limit ? files.slice(0, limit) : files;
  if (!selected.length) throw new Error('没有匹配的 xlsx 文件');
  return selected;
}

function findExcelUrl(value) {
  if (typeof value === 'string') return /^https?:\/\/.*\.xlsx(?:\?|$)/i.test(value) ? value : '';
  if (Array.isArray(value)) {
    for (const item of value) {
      const found = findExcelUrl(item);
      if (found) return found;
    }
  } else if (value && typeof value === 'object') {
    for (const item of Object.values(value)) {
      const found = findExcelUrl(item);
      if (found) return found;
    }
  }
  return '';
}

async function uploadExcel(baseUrl, filePath) {
  const name = path.basename(filePath);
  const form = new FormData();
  form.append('type', 'OTHER');
  form.append('part', new Blob([fs.readFileSync(filePath)], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  }), name);
  form.append('title', name);
  const response = await axios.post(`${baseUrl}/files`, form, {
    headers: getHeaders(), timeout: 120000, maxBodyLength: Infinity,
  });
  if (response.data?.code !== 0) throw new Error(response.data?.message || '文件上传失败');
  const excelUrl = findExcelUrl(response.data);
  if (!excelUrl) throw new Error('上传响应中没有可识别的 xlsx URL');
  return { excelUrl, response: response.data };
}

async function submitBatch(baseUrl, excelUrl, options) {
  const response = await axios.post(`${baseUrl}/batch/update/place/welmsg`, {
    placeType: options.placeType,
    excelUrl,
    placeName: options.placeName,
  }, { headers: getHeaders('application/json;charset=UTF-8'), timeout: 120000 });
  if (response.data?.code !== 0) throw new Error(response.data?.message || '批量导入失败');
  return response.data;
}

function writeReport(rows) {
  const timestamp = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14);
  const outputDir = path.resolve('output');
  fs.mkdirSync(outputDir, { recursive: true });
  const outputPath = path.join(outputDir, `渠道欢迎语批量导入_${timestamp}.json`);
  fs.writeFileSync(outputPath, `${JSON.stringify(rows, null, 2)}\n`, 'utf8');
  return outputPath;
}

async function main() {
  const options = parseOptions();
  if (options.execute && !options.uploadOnly && !options.backupConfirmed) {
    throw new Error('正式导入前必须完成原始数据备份，并追加 --backup-confirmed');
  }
  const baseUrl = `https://vinci-api.feibing.tech/sc/v1/sellers/${options.sellerId}`;
  const files = options.excelUrl ? [] : listExcelFiles(options.inputDir, options.start, options.limit);
  console.log(JSON.stringify({
    mode: options.execute ? (options.uploadOnly ? 'upload-only' : 'execute') : 'dry-run',
    sellerId: options.sellerId,
    placeName: options.placeName,
    placeType: options.placeType,
    files: files.map(file => path.basename(file)),
    excelUrl: options.excelUrl || undefined,
  }, null, 2));
  if (!options.execute) return;

  const work = options.excelUrl ? [{ excelUrl: options.excelUrl }] : files.map(file => ({ file }));
  const report = [];
  for (const source of work) {
    const item = { file: source.file ? path.basename(source.file) : '', status: 'pending' };
    report.push(item);
    try {
      if (source.excelUrl) {
        item.excelUrl = source.excelUrl;
      } else {
        const uploaded = await uploadExcel(baseUrl, source.file);
        item.excelUrl = uploaded.excelUrl;
        item.uploadResponse = uploaded.response;
        item.status = 'uploaded';
      }
      if (!options.uploadOnly) {
        item.submitResponse = await submitBatch(baseUrl, item.excelUrl, options);
        item.status = 'submitted';
      }
    } catch (error) {
      item.status = 'failed';
      item.error = error.response?.data || error.message;
      writeReport(report);
      throw error;
    }
  }
  console.log(`处理完成：${writeReport(report)}`);
}

if (require.main === module) {
  main().catch(error => {
    console.error(error.response?.data || error.message);
    process.exitCode = 1;
  });
}

module.exports = { findExcelUrl, listExcelFiles, parseOptions };
