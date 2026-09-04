/**
 * @对象    通用
 * @动作    导出
 * @风险    低
 * @来源    store-data-extractor/DuQuShuJu.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

// 导入依赖
const axios = require('axios');
const ExcelJS = require('exceljs');
const fs = require('fs-extra');
const path = require('path');

// ===================== 核心配置（需确认/替换）=====================
const API_URL = 'https://vinci-api.feibing.tech/sc/v1/sellers/wwa9c5a585540b115b/vinci_shop_rpt_fb_customer_stat/statistics';
const AUTH_TOKEN = '__VINCI_TOKEN__';
const STAT_DATE = '2025-09-30'; // 查询日期（按需修改）
const PAGE_SIZE = 100; // 每页数据量（固定100）
const BATCH_PAGE_COUNT = 100; // 每100页生成一个Excel文件
const OUTPUT_DIR = path.join(__dirname, '../output/拉新数据'); // Excel输出目录

// 请求头（从curl提取，已包含必要字段）
const REQUEST_HEADERS = {
  'accept': 'application/json',
  'accept-language': 'zh-CN,zh;q=0.9',
  'authorization': `Bearer ${AUTH_TOKEN}`,
  'content-type': 'application/json;charset=UTF-8',
  'origin': 'http://172.20.10.4:1024',
  'referer': 'http://172.20.10.4:1024/',
  'sec-ch-ua': '"Chromium";v="140", "Not=A?Brand";v="24", "Google Chrome";v="140"',
  'sec-ch-ua-mobile': '?0',
  'sec-ch-ua-platform': '"macOS"',
  'sec-fetch-dest': 'empty',
  'sec-fetch-mode': 'cors',
  'sec-fetch-site': 'cross-site',
  'sec-fetch-storage-access': 'active',
  'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36'
};

// ===================== 工具函数 =====================
/**
 * 单个页面数据请求
 * @param {number} pageNumber - 当前页码（0开始）
 * @returns {Promise<{list: Array, total: number}>} - 页面数据列表 + 总数据量
 */
async function fetchPageData(pageNumber) {
  // 构建请求体（提取为变量，方便重试时复用）
  const requestBody = {
    ext: {
      query: {
        page: {
          pageNumber: pageNumber,
          pageSize: PAGE_SIZE,
          orders: [
            { field: 'lx_all_nums', direction: 'DESC' },
            { field: 'shop_eid', direction: 'DESC' }
          ]
        }
      }
    },
    statDate: STAT_DATE,
    shopStatus: 'PUBLISHED'
  };

  try {
    const response = await axios.post(API_URL, requestBody, { headers: REQUEST_HEADERS });

    // 解析返回数据（注意：data字段是字符串，需JSON转义）
    const result = JSON.parse(response.data.data);
    if (!result.list || !result.total) {
      throw new Error(`第${pageNumber+1}页数据格式异常，无list/total字段`);
    }
    console.log(`✅ 成功获取第${pageNumber+1}页数据，共${result.list.length}条`);
    return {
      list: result.list,
      total: result.total
    };
  } catch (error) {
    console.error(`❌ 第${pageNumber+1}页请求失败：`, error.message);
    // 重试机制：失败后重试2次（避免网络波动）
    for (let retry = 1; retry <= 2; retry++) {
      console.log(`🔄 第${retry}次重试第${pageNumber+1}页`);
      try {
        const response = await axios.post(API_URL, requestBody, { headers: REQUEST_HEADERS });
        const result = JSON.parse(response.data.data);
        console.log(`✅ 重试成功：第${pageNumber+1}页`);
        return { list: result.list, total: result.total };
      } catch (retryError) {
        console.error(`❌ 第${retry}次重试失败：`, retryError.message);
      }
    }
    throw new Error(`第${pageNumber+1}页请求失败（已重试2次）`);
  }
}

/**
 * 收集所有页面的完整数据（新增“店铺id”“店铺三方Code”字段）
 * @returns {Promise<Array>} - 所有店铺的处理后数据（七列结构）
 */
async function collectAllData() {
  console.log('📥 开始收集所有页面数据...');
  // 1. 先请求第1页（pageNumber=0），获取总数据量和总页数
  const firstPage = await fetchPageData(0);
  const totalDataCount = firstPage.total; // 总数据量
  const totalPages = Math.ceil(totalDataCount / PAGE_SIZE); // 总页数
  console.log(`📊 数据概况：总数据量${totalDataCount}条，总页数${totalPages}页`);

  // 2. 处理第1页数据（新增“店铺id”“店铺三方Code”，对应shop_eid、third_shop_code）
  const allData = firstPage.list.map(shop => ({
    shopId: shop.shop_eid || '', // 新增：店铺id（读取shop_eid）
    thirdShopCode: shop.third_shop_code || '', // 新增：店铺三方Code（读取third_shop_code）
    shopName: shop.shop_name || '', // 原有：店铺名称
    groupNums: shop.group_nums || 0, // 原有：群聊数
    groupUsers: shop.group_users || 0, // 原有：群聊客户总人数
    lostUsers: shop.group_customers_lost || 0, // 原有：客户流失数
    retainedUsers: (shop.group_users || 0) - (shop.group_customers_lost || 0) // 原有：留存人数（计算值）
  }));

  // 3. 循环请求剩余页面（同样新增两列数据）
  for (let pageNumber = 1; pageNumber < totalPages; pageNumber++) {
    const pageData = await fetchPageData(pageNumber);
    const formattedPageData = pageData.list.map(shop => ({
      shopId: shop.shop_eid || '', // 新增：店铺id
      thirdShopCode: shop.third_shop_code || '', // 新增：店铺三方Code
      shopName: shop.shop_name || '',
      groupNums: shop.group_nums || 0,
      groupUsers: shop.group_users || 0,
      lostUsers: shop.group_customers_lost || 0,
      retainedUsers: (shop.group_users || 0) - (shop.group_customers_lost || 0)
    }));
    allData.push(...formattedPageData);
  }

  console.log(`📥 所有数据收集完成，共${allData.length}条店铺记录（含新增列）`);
  return allData;
}

/**
 * 生成Excel文件（批次文件+汇总文件，均新增两列表头）
 * @param {Array} allData - 所有店铺的处理后数据（七列结构）
 */
async function generateExcelFiles(allData) {
  console.log('📋 开始生成Excel文件...');
  // 1. 创建输出目录（不存在则自动创建）
  await fs.ensureDir(OUTPUT_DIR);

  // 2. 计算批次数量
  const batchDataSize = PAGE_SIZE * BATCH_PAGE_COUNT;
  const totalBatches = Math.ceil(allData.length / batchDataSize);

  // 3. 按批次生成Excel文件（表头新增“店铺id”“店铺三方Code”）
  for (let batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
    const startIndex = batchIndex * batchDataSize;
    const endIndex = Math.min((batchIndex + 1) * batchDataSize, allData.length);
    const batchData = allData.slice(startIndex, endIndex);
    const startPage = batchIndex * BATCH_PAGE_COUNT + 1;
    const endPage = Math.min((batchIndex + 1) * BATCH_PAGE_COUNT, Math.ceil(allData.length / PAGE_SIZE));
    const fileName = `店铺数据_${startPage}-${endPage}页.xlsx`;
    const filePath = path.join(OUTPUT_DIR, fileName);

    // 创建工作簿和工作表
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(`数据_${startPage}-${endPage}页`);

    // 设置表头（新增两列，调整宽度适配内容）
    worksheet.columns = [
      { header: '店铺id', key: 'shopId', width: 30 }, // 新增：店铺id（对应数据字段shopId）
      { header: '店铺三方Code', key: 'thirdShopCode', width: 20 }, // 新增：店铺三方Code（对应thirdShopCode）
      { header: '店铺名称', key: 'shopName', width: 30 }, // 原有列
      { header: '群聊数', key: 'groupNums', width: 12 }, // 原有列
      { header: '群聊客户总人数', key: 'groupUsers', width: 18 }, // 原有列
      { header: '客户流失数', key: 'lostUsers', width: 12 }, // 原有列
      { header: '留存人数', key: 'retainedUsers', width: 12 } // 原有列
    ];

    // 添加数据到工作表
    worksheet.addRows(batchData);
    // 保存批次文件
    await workbook.xlsx.writeFile(filePath);
    console.log(`✅ 批次文件生成完成：${fileName}（${batchData.length}条数据，含新增列）`);
  }

  // 4. 生成汇总Excel文件（表头同样新增两列）
  const summaryFileName = `店铺数据_汇总.xlsx`;
  const summaryFilePath = path.join(OUTPUT_DIR, summaryFileName);
  const summaryWorkbook = new ExcelJS.Workbook();
  const summaryWorksheet = summaryWorkbook.addWorksheet('所有店铺数据汇总');

  // 汇总表表头（与批次表一致，含新增列）
  summaryWorksheet.columns = [
    { header: '店铺id', key: 'shopId', width: 30 },
    { header: '店铺三方Code', key: 'thirdShopCode', width: 20 },
    { header: '店铺名称', key: 'shopName', width: 30 },
    { header: '群聊数', key: 'groupNums', width: 12 },
    { header: '群聊客户总人数', key: 'groupUsers', width: 18 },
    { header: '客户流失数', key: 'lostUsers', width: 12 },
    { header: '留存人数', key: 'retainedUsers', width: 12 }
  ];

  // 添加所有数据到汇总表
  summaryWorksheet.addRows(allData);
  // 保存汇总文件
  await summaryWorkbook.xlsx.writeFile(summaryFilePath);
  console.log(`✅ 汇总文件生成完成：${summaryFileName}（${allData.length}条数据，含新增列）`);

  console.log(`\n🎉 所有Excel文件生成完成！输出目录：${OUTPUT_DIR}`);
}

// ===================== 主执行函数 =====================
async function main() {
  try {
    const allData = await collectAllData();
    await generateExcelFiles(allData);
    console.log('✅ 任务全部完成！Excel已新增“店铺id”“店铺三方Code”列');
  } catch (error) {
    console.error('❌ 任务执行失败：', error.message);
  }
}

// 启动脚本
main();