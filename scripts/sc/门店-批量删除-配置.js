/**
 * @对象    门店
 * @动作    批量写
 * @风险    高
 * @来源    store-data-extractor/QuanMenDianShanChuPeiZhi.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const fs = require('fs').promises;
const path = require('path');
const xlsx = require('xlsx');
const axios = require('axios');
const { existsSync, mkdirSync } = require('fs');

// 根据表格里面的门店id 去删除对应的配置
// 1. 读取Excel，提取“店铺id”
// 2. 调用GET接口，获取对应配置的id
// 3. 调用DELETE接口，删除对应配置

// -------------------------- 配置项（根据实际情况调整）--------------------------
const CONFIG = {
  EXCEL_PATH: path.join(__dirname, '../input/全门店-沪上.xlsx'), // Excel文件路径
  EXCEL_SHEET_NAME: 'Sheet1', // Excel工作表名
  EXCEL_SHOP_ID_COLUMN: '店铺id', // Excel中“店铺id”的列名（务必与Excel一致）
  GET_API_URL: 'https://vinci-api.feibing.tech/sc/v1/sellers/wwa9c5a585540b115b/configs', // GET接口基础URL
  DELETE_API_BASE_URL: 'https://vinci-api.feibing.tech/sc/v1/sellers/wwa9c5a585540b115b/configs', // DELETE接口基础URL
  CONCURRENT_NUM: 3, // 并发数（一次调用5个）
  OUTPUT_LOG_PATH: path.join(__dirname, '../output/delete_result.log'), // 结果日志路径
  // 请求头（注意：Bearer Token可能过期，需手动更新！）
  REQUEST_HEADERS: {
    'accept': '*/*',
    'accept-language': 'zh-CN,zh;q=0.9',
    'authorization': 'bearer __VINCI_TOKEN__',
    'cookie': '_clck=sj51u%5E2%5Efyn%5E0%5E1974; x-token=__VINCI_TOKEN__',
    'origin': 'https://connect.feibing.tech',
    'priority': 'u=1, i',
    'referer': 'https://connect.feibing.tech/',
    'sec-ch-ua': '"Not;A=Brand";v="99", "Google Chrome";v="139", "Chromium";v="139"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"macOS"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-site',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36'
  }
};
// -------------------------------------------------------------------------------

// 初始化：创建output文件夹（日志存放目录）
async function initOutputFolder() {
  const outputDir = path.dirname(CONFIG.OUTPUT_LOG_PATH);
  if (!existsSync(outputDir)) {
    mkdirSync(outputDir, { recursive: true });
    console.log(`✅ 已创建输出文件夹：${outputDir}`);
  }
}

// 读取Excel，提取“店铺id”和对应行号（Excel行号从2开始，对应数据行）
async function readShopIdsFromExcel() {
  try {
    if (!existsSync(CONFIG.EXCEL_PATH)) {
      throw new Error(`Excel文件不存在：${CONFIG.EXCEL_PATH}`);
    }

    const workbook = xlsx.readFile(CONFIG.EXCEL_PATH);
    const worksheet = workbook.Sheets[CONFIG.EXCEL_SHEET_NAME];
    if (!worksheet) {
      throw new Error(`Excel工作表不存在：${CONFIG.EXCEL_SHEET_NAME}`);
    }

    const jsonData = xlsx.utils.sheet_to_json(worksheet);
    if (jsonData.length === 0) {
      throw new Error('Excel中无数据');
    }

    const firstRowKeys = Object.keys(jsonData[0]);
    if (!firstRowKeys.includes(CONFIG.EXCEL_SHOP_ID_COLUMN)) {
      throw new Error(`Excel中未找到“${CONFIG.EXCEL_SHOP_ID_COLUMN}”列，现有列名：${firstRowKeys.join(', ')}`);
    }

    const shopList = jsonData.map((row, index) => ({
      shopId: row[CONFIG.EXCEL_SHOP_ID_COLUMN]?.toString().trim(),
      excelRow: index + 2
    })).filter(item => item.shopId);

    console.log(`✅ 从Excel读取到 ${shopList.length} 个有效店铺id`);
    return shopList;
  } catch (err) {
    console.error(`❌ 读取Excel失败：${err.message}`);
    process.exit(1);
  }
}

// 记录日志到文件
async function logResult(logInfo) {
  const timestamp = new Date().toLocaleString();
  const logContent = `[${timestamp}] | Excel行${logInfo.excelRow} | 店铺id:${logInfo.shopId} | 状态:${logInfo.status} | 原因:${logInfo.reason || '无'}\n`;
  
  try {
    await fs.appendFile(CONFIG.OUTPUT_LOG_PATH, logContent);
  } catch (err) {
    console.error(`❌ 写入日志失败：${err.message}，日志内容：${logContent}`);
  }
}

// 调用GET接口：新增“返回空数组时跳过DELETE”逻辑
async function callGetConfigApi(shopItem) {
  try {
    const key = `${shopItem.shopId}_pre_h5_config_wx`;
    const getUrl = `${CONFIG.GET_API_URL}?type=MINI_APP_STYLE_TYPE&key=${encodeURIComponent(key)}`;

    const getResponse = await axios({
      method: 'get',
      url: getUrl,
      headers: CONFIG.REQUEST_HEADERS,
      timeout: 15000
    });

    const getResult = getResponse.data;
    // 核心判断：code=0但total=0或data为空数组 → 无需删除，返回null
    if (getResult.code === 0 && (getResult.total === 0 || !getResult.data || getResult.data.length === 0)) {
      console.log(`📥 Excel行${shopItem.excelRow} | 店铺id${shopItem.shopId} | GET成功但返回空数组（total=0/data=[]），无需执行DELETE`);
      return null; // 返回null标识“无需删除”
    }
    // 正常有数据的情况
    if (getResult.code !== 0 || !getResult.data || getResult.data.length === 0) {
      throw new Error(`GET接口返回异常：code=${getResult.code || '无'}, message=${getResult.message || '无数据'}`);
    }

    const configId = getResult.data[0].id;
    console.log(`📥 Excel行${shopItem.excelRow} | 店铺id${shopItem.shopId} | GET成功，获取到configId：${configId}`);
    return configId;
  } catch (err) {
    const errorMsg = err.response 
      ? `HTTP${err.response.status}：${err.response.data?.message || '接口错误'}` 
      : err.message;
    throw new Error(`GET接口失败：${errorMsg}`);
  }
}

// 调用DELETE接口
async function callDeleteConfigApi(shopItem, configId) {
  try {
    const deleteUrl = `${CONFIG.DELETE_API_BASE_URL}/${encodeURIComponent(configId)}`;

    await axios({
      method: 'delete',
      url: deleteUrl,
      headers: CONFIG.REQUEST_HEADERS,
      timeout: 15000
    });

    console.log(`🗑️ Excel行${shopItem.excelRow} | 店铺id${shopItem.shopId} | DELETE成功，configId：${configId}`);
    return true;
  } catch (err) {
    const errorMsg = err.response 
      ? `HTTP${err.response.status}：${err.response.data?.message || '接口错误'}` 
      : err.message;
    throw new Error(`DELETE接口失败：${errorMsg}`);
  }
}

// 单个店铺处理逻辑：新增“GET返回null时跳过DELETE”
async function processSingleShop(shopItem) {
  try {
    const configId = await callGetConfigApi(shopItem);
    // 核心判断：configId为null → 跳过DELETE，记录“无需删除”日志
    if (configId === null) {
      await logResult({
        excelRow: shopItem.excelRow,
        shopId: shopItem.shopId,
        status: '无需删除',
        reason: 'GET接口返回空数组（total=0/data=[]），无配置需删除'
      });
      return; // 直接结束，不执行DELETE
    }
    // 有configId则执行DELETE
    await callDeleteConfigApi(shopItem, configId);
    await logResult({
      excelRow: shopItem.excelRow,
      shopId: shopItem.shopId,
      status: '成功',
      reason: `GET+DELETE均成功，configId：${configId}`
    });
  } catch (err) {
    await logResult({
      excelRow: shopItem.excelRow,
      shopId: shopItem.shopId,
      status: '失败',
      reason: err.message
    });
    console.error(`⚠️ Excel行${shopItem.excelRow} | 店铺id${shopItem.shopId} | 处理失败：${err.message}`);
  }
}

// 分批并发处理（逻辑不变，保持5个店铺并发）
async function batchProcessShops(shopList) {
  const totalCount = shopList.length;
  const batchSize = CONFIG.CONCURRENT_NUM;
  const totalBatches = Math.ceil(totalCount / batchSize);

  console.log(`\n✅ 开始分批处理：共${totalCount}个店铺，每批${batchSize}个，共${totalBatches}批`);

  for (let batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
    const start = batchIndex * batchSize;
    const end = Math.min(start + batchSize, totalCount);
    const currentBatch = shopList.slice(start, end);
    const batchNum = batchIndex + 1;

    console.log(`\n🚀 处理第${batchNum}/${totalBatches}批：店铺${start+1}-${end}/${totalCount}`);

    const batchPromises = currentBatch.map(shopItem => processSingleShop(shopItem));
    await Promise.allSettled(batchPromises);

    console.log(`✅ 第${batchNum}/${totalBatches}批处理完成`);
  }
}

// 主函数
async function main() {
  console.log('========================================');
  console.log('🔍 开始执行店铺配置删除脚本');
  console.log('========================================');

  try {
    await initOutputFolder();
    const shopList = await readShopIdsFromExcel();
    if (shopList.length === 0) {
      console.log('❌ 无有效店铺id，脚本终止');
      return;
    }
    await batchProcessShops(shopList);
    console.log('\n========================================');
    console.log(`🎉 所有店铺处理完成！结果日志：${CONFIG.OUTPUT_LOG_PATH}`);
    console.log('========================================');
  } catch (err) {
    console.error(`\n❌ 脚本执行异常：${err.message}`);
  }
}

// 启动脚本
main();