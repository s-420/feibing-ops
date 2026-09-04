/**
 * @对象    门店
 * @动作    批量写
 * @风险    高
 * @来源    store-data-extractor/PiLiangTongBuMenDianPeiZhi.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const axios = require('axios');
const XLSX = require('xlsx');
const path = require('path');

// ========================== 1. 基础配置（复用鉴权，统一维护）==========================
// 通用Headers（所有接口复用，避免重复）
const COMMON_HEADERS = {
  'accept': '*/*',
  'accept-language': 'zh-CN,zh;q=0.9',
  'authorization': 'bearer __VINCI_TOKEN__',
  'cookie': '_clck=sj51u%5E2%5Efyn%5E0%5E1974; x-token=__VINCI_TOKEN__; acw_tc=0bca30f217578648974194519e3c5696d3d252502fd57900832563191bf95e',
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
};

// 接口URL模板（动态参数用{xxx}占位）
const API_TEMPLATES = {
  // 1. 获取默认value的接口（固定）
  getDefaultValue: 'https://vinci-api.feibing.tech/sc/v1/sellers/wpV3cNDAAA3ZRQDQZt0UsR9wORkx0thg/configs?type=MINI_APP_STYLE_TYPE&key=69b3b942c3242c368fa302c6_pre_h5_config_wx',
  // 2. 查询门店配置ID的接口（key=门店ID_pre_h5_config_wx）
  getStoreConfigId: 'https://vinci-api.feibing.tech/sc/v1/sellers/wwa9c5a585540b115b/configs?type=MINI_APP_STYLE_TYPE&key={storeKey}',
  // 3. 更新门店配置的接口（末尾ID=查询到的配置ID）
  updateStoreConfig: 'https://vinci-api.feibing.tech/sc/v1/sellers/wwa9c5a585540b115b/configs/{configId}'
};

// Excel配置
const EXCEL_CONFIG = {
  path: path.join(__dirname, '../input/同步企业设置.xlsx'), // 根目录input表格
  storeIdColumnName: '门店ID' // 确认Excel列名
};

// 请求延迟（避免限流）
const REQUEST_DELAY = 500;


// ========================== 2. 核心函数（新增“查询配置ID”步骤）==========================
/**
 * 函数1：调用第一个接口，获取默认value（不变）
 */
async function getDefaultValue() {
  try {
    console.log('✅ 【步骤1/3】调用默认配置接口，获取基础value...');
    const response = await axios.get(API_TEMPLATES.getDefaultValue, { headers: COMMON_HEADERS });

    const { code, data, message } = response.data;
    if (code !== 0) throw new Error(`默认接口失败：${message}（code: ${code}）`);
    if (!data || data.length === 0) throw new Error('默认接口返回无数据');

    const defaultValue = data[0].value;
    if (!defaultValue) throw new Error('默认接口的value为空');

    console.log('✅ 【步骤1/3】默认value获取成功（前80字符）：', defaultValue.slice(0, 80) + '...\n');
    return defaultValue;
  } catch (error) {
    console.error('❌ 【步骤1/3】默认接口调用失败：', error.message);
    throw error; // 终止流程（无基础value无法继续）
  }
}

/**
 * 函数2：读取Excel中的门店ID（不变）
 */
async function readStoreIdsFromExcel() {
  try {
    console.log('✅ 【步骤2/3】读取Excel中的门店ID列表...');
    const workbook = XLSX.readFile(EXCEL_CONFIG.path);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    if (excelData.length <= 1) throw new Error('Excel仅含表头或无数据');
    const headerRow = excelData[0];
    const storeIdIndex = headerRow.findIndex(col => col === EXCEL_CONFIG.storeIdColumnName);
    if (storeIdIndex === -1) throw new Error(`Excel无“${EXCEL_CONFIG.storeIdColumnName}”列`);

    // 提取有效门店ID（过滤空值）
    const storeIds = excelData.slice(1)
      .map(row => row[storeIdIndex])
      .filter(id => id && String(id).trim() !== '');

    if (storeIds.length === 0) throw new Error('Excel未提取到有效门店ID');
    console.log(`✅ 【步骤2/3】Excel读取成功！共${storeIds.length}个门店：`, storeIds.slice(0, 5) + (storeIds.length > 5 ? '...' : '') + '\n');
    return storeIds;
  } catch (error) {
    console.error('❌ 【步骤2/3】Excel读取失败：', error.message);
    throw error;
  }
}

/**
 * 函数3：新增！查询单个门店的配置ID（从用户提供的查询接口获取）
 * @param {string} storeId  Excel中的门店ID
 * @returns {Promise<string>} 门店对应的配置ID（如68bd9b97b9fa2b7969875fd8）
 */
async function getStoreConfigId(storeId) {
  try {
    // 1. 构造查询接口的key（门店ID + 后缀）
    const storeKey = `${storeId}_pre_h5_config_wx`;
    // 2. 替换URL模板中的{storeKey}，生成查询URL
    const requestUrl = API_TEMPLATES.getStoreConfigId.replace('{storeKey}', storeKey);

    console.log(`🔍 【门店${storeId}】查询配置ID：key=${storeKey}`);
    console.log(`   查询接口URL：${requestUrl}`);
    const response = await axios.get(requestUrl, { headers: COMMON_HEADERS });

    // 3. 解析响应，提取配置ID（data[0].id）
    const { code, data, message } = response.data;
    if (code !== 0) throw new Error(`查询失败：${message}（code: ${code}）`);
    if (!data || data.length === 0) throw new Error('查询返回无配置数据');
    if (!data[0].id) throw new Error('返回数据中无id字段');

    const configId = data[0].id;
    console.log(`✅ 【门店${storeId}】查询到配置ID：${configId}\n`);
    return configId;
  } catch (error) {
    console.error(`❌ 【门店${storeId}】配置ID查询失败：`, error.message);
    return null; // 返回null表示查询失败，后续跳过该门店
  }
}

/**
 * 函数4：更新单个门店配置（先查ID，再更新）
 * @param {string} storeId  门店ID
 * @param {string} defaultValue  默认value
 * @returns {Promise<boolean>} 更新成功返回true
 */
async function updateSingleStoreConfig(storeId, defaultValue) {
  // 步骤1：先查询该门店的配置ID
  const configId = await getStoreConfigId(storeId);
  if (!configId) {
    console.warn(`⚠️  【门店${storeId}】跳过更新（无有效配置ID）\n`);
    return false;
  }

  // 步骤2：用配置ID构造更新接口URL，调用PUT请求
  try {
    const requestUrl = API_TEMPLATES.updateStoreConfig.replace('{configId}', configId);
    const requestBody = {
      type: 'MINI_APP_STYLE_TYPE', // 固定值
      key: `${storeId}_pre_h5_config_wx`, // 与查询接口的key一致
      value: defaultValue // 复用默认value
    };

    console.log(`🚀 【门店${storeId}】开始更新配置（ID: ${configId}）`);
    const response = await axios.put(requestUrl, requestBody, { headers: COMMON_HEADERS });

    // 校验更新结果
    const { code, message } = response.data;
    if (code !== 0) throw new Error(`更新失败：${message}（code: ${code}）`);

    console.log(`✅ 【门店${storeId}】配置更新成功！\n`);
    return true;
  } catch (error) {
    console.error(`❌ 【门店${storeId}】配置更新失败：`, error.message);
    return false;
  }
}

/**
 * 延迟函数（避免接口限流）
 */
function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * 主函数：串联完整流程
 */
async function main() {
  try {
    console.log('=====================================================');
    console.log('🚀 批量同步门店配置脚本（最终版） 开始执行');
    console.log('=====================================================\n');

    // 阶段1：获取默认value + 读取门店ID
    const defaultValue = await getDefaultValue();
    const storeIds = await readStoreIdsFromExcel();

    // 阶段2：循环处理每个门店（查ID→更配置）
    let successCount = 0;
    console.log('✅ 【步骤3/3】开始批量处理门店配置...\n');
    for (let i = 0; i < storeIds.length; i++) {
      const storeId = storeIds[i];
      console.log(`=====================================================`);
      console.log(`📌 正在处理第 ${i+1}/${storeIds.length} 个门店：${storeId}`);
      console.log(`=====================================================`);

      // 更新门店配置
      const isSuccess = await updateSingleStoreConfig(storeId, defaultValue);
      if (isSuccess) successCount++;

      // 最后一个门店不延迟
      if (i < storeIds.length - 1) {
        console.log(`⌛ 等待 ${REQUEST_DELAY/1000} 秒后处理下一个门店...\n`);
        await delay(REQUEST_DELAY);
      }
    }

    // 输出最终结果
    console.log('=====================================================');
    console.log('📊 脚本执行完成！');
    console.log(`总门店数：${storeIds.length} | 成功数：${successCount} | 失败数：${storeIds.length - successCount}`);
    console.log('=====================================================');
  } catch (error) {
    console.error('\n❌ 脚本执行异常终止（基础步骤失败）：', error.message);
    console.log('=====================================================');
  }
}

// 启动脚本
main();