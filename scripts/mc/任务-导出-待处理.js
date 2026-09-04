/**
 * @对象    任务
 * @动作    导出
 * @风险    低
 * @来源    store-data-extractor/DaoChuTaskExecutorsPending.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const axios = require('axios');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// 导出旧版星图朋友圈任务数据   青岛啤酒企业
const HARDCODED_AUTHORIZATION =
  'bearer __VINCI_TOKEN__';

const config = {
  baseUrl: 'https://vinci-api.feibing.tech/mc/v1/task/executors',
  pageSize: Number(process.env.PAGE_SIZE || 10),
  sellerId: process.env.SELLER_ID || 'wpV3cNDAAAa5pLikCaI4KbhtYUxy9whw',
  staffWxUserid: process.env.STAFF_WX_USERID || 'woV3cNDAAA8wsSx1kxrrdwH85qfV-OCw',
  status: process.env.TASK_STATUS || 'PENDING',
  authorization: process.env.AUTHORIZATION || HARDCODED_AUTHORIZATION,
  outputDir: path.join(__dirname, '../task_executors_excels'),
  outputFileName:
    process.env.OUTPUT_FILE || `task_executors_${Date.now()}.xlsx`
};

function ensureOutputDir() {
  if (!fs.existsSync(config.outputDir)) {
    fs.mkdirSync(config.outputDir, { recursive: true });
  }
}

function formatTimestamp(timestamp) {
  if (!timestamp) return '';
  const date = new Date(Number(timestamp));
  if (Number.isNaN(date.getTime())) return '';
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  const hh = String(date.getHours()).padStart(2, '0');
  const mi = String(date.getMinutes()).padStart(2, '0');
  const ss = String(date.getSeconds()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd} ${hh}:${mi}:${ss}`;
}

function parseTaskMeta(item) {
  const taskRaw = item?.metadata?.task;
  if (!taskRaw || typeof taskRaw !== 'string') {
    return { done: '', pending: '' };
  }

  try {
    const taskObj = JSON.parse(taskRaw);
    return {
      done: taskObj?.DONE ?? '',
      pending: taskObj?.PENDING ?? ''
    };
  } catch (error) {
    return { done: '', pending: '' };
  }
}

async function fetchOnePage(pageNum) {
  const response = await axios.get(config.baseUrl, {
    params: {
      staffWxUserid: config.staffWxUserid,
      pageSize: config.pageSize,
      pageNum,
      sellerId: config.sellerId,
      status: config.status
    },
    headers: {
      Authorization: config.authorization,
      Accept: '*/*',
      'content-type': 'application/json'
    },
    timeout: 30000
  });

  const body = response.data || {};
  if (body.code !== 0) {
    throw new Error(`接口返回异常: code=${body.code}, message=${body.message}`);
  }
  return body;
}

async function fetchAllData() {
  const firstPage = await fetchOnePage(1);
  const total = Number(firstPage.total || 0);
  const totalPages = Math.max(1, Math.ceil(total / config.pageSize));
  let allRows = Array.isArray(firstPage.data) ? [...firstPage.data] : [];

  console.log(`总条数: ${total}, 总页数: ${totalPages}`);

  for (let pageNum = 2; pageNum <= totalPages; pageNum++) {
    console.log(`拉取第 ${pageNum}/${totalPages} 页...`);
    const pageResult = await fetchOnePage(pageNum);
    const pageRows = Array.isArray(pageResult.data) ? pageResult.data : [];
    allRows = allRows.concat(pageRows);
  }

  return allRows;
}

async function exportExcel(rows) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('task_executors');

  sheet.columns = [
    { header: 'taskName', key: 'taskName', width: 80 },
    { header: 'DONE', key: 'done', width: 12 },
    { header: 'PENDING', key: 'pending', width: 12 },
    { header: 'appMsgTime', key: 'appMsgTime', width: 16 },
    { header: 'appMsgTimeFormatted', key: 'appMsgTimeFormatted', width: 22 }
  ];

  rows.forEach(item => {
    const taskName = item?.task?.name || '';
    const appMsgTime = item?.appMsgTime ?? '';
    const { done, pending } = parseTaskMeta(item);

    sheet.addRow({
      taskName,
      done,
      pending,
      appMsgTime,
      appMsgTimeFormatted: formatTimestamp(appMsgTime)
    });
  });

  const outputPath = path.join(config.outputDir, config.outputFileName);
  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}

async function main() {
  try {
    ensureOutputDir();
    console.log('开始导出 task executors...');
    const allRows = await fetchAllData();
    const outputPath = await exportExcel(allRows);
    console.log(`导出完成，共 ${allRows.length} 条`);
    console.log(`文件路径: ${outputPath}`);
  } catch (error) {
    console.error(`执行失败: ${error.message}`);
    process.exit(1);
  }
}

main();
