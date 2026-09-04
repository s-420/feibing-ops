/**
 * @对象    客户
 * @动作    导出
 * @风险    低
 * @来源    store-data-extractor/DaoChuKeHuEidToUidXinXi.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const axios = require('axios');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 配置项
const CONFIG = {
  // 输入Excel路径
  inputExcelPath: path.resolve(__dirname, '../input/客户Eid.xlsx'),
  // 输出目录
  outputDir: path.resolve(__dirname, '../output'),
  // 输出Excel文件名
  outputFileName: '客户Eid_UID.xlsx',
  // API基础URL
  apiBaseUrl: 'https://vinci-api.feibing.tech/sc/v1/ydd/eid2uid',
  // 请求头（从curl复制转换）
  headers: {
    'Accept-Encoding': 'gzip, deflate, br',
    'Connection': 'keep-alive',
    'Cookie': 'acw_tc=0bca31d217627624787992881ea492b33a7163b016f6b780ca9662e6ae91a3',
    'accept': '*/*',
    'accept-language': 'zh-CN,zh;q=0.9',
    'authorization': 'bearer __VINCI_TOKEN__',
    'origin': 'https://connect.feibing.tech',
    'priority': 'u=1, i',
    'referer': 'https://connect.feibing.tech/',
    'sec-ch-ua': '"Not)A;Brand";v="8", "Chromium";v="138", "Google Chrome";v="138"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"macOS"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-site',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36'
  },
  // 并发批次大小（一次10条）
  batchSize: 10,
  // 批次之间的延迟（毫秒），避免请求过于密集
  batchDelay: 500
};

/**
 * 读取输入Excel中的name和externalUserId
 * @returns {Array<{name: string, eid: string}>} 包含姓名和EID的数组
 */
function readCustomerDataFromExcel() {
  try {
    console.log(`正在读取输入文件: ${CONFIG.inputExcelPath}`);
    // 读取Excel文件
    const workbook = XLSX.readFile(CONFIG.inputExcelPath);
    // 获取第一个工作表
    const worksheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[worksheetName];
    // 转换为JSON
    const data = XLSX.utils.sheet_to_json(worksheet);

    // 提取name和externalUserId字段
    const customerList = data
      .map(item => ({
        name: item.name || '未知姓名',
        eid: item.externalUserId
      }))
      .filter(item => item.eid && typeof item.eid === 'string'); // 过滤无效EID

    if (customerList.length === 0) {
      throw new Error('未在Excel中找到有效的externalUserId数据');
    }

    console.log(`成功读取到 ${customerList.length} 条客户数据`);
    return customerList;
  } catch (error) {
    console.error('读取Excel失败:', error.message);
    process.exit(1);
  }
}

/**
 * 调用API获取EID对应的UID
 * @param {string} eid - 客户EID
 * @returns {Promise<string>} UID或错误信息
 */
async function fetchUidByEid(eid) {
  try {
    const url = `${CONFIG.apiBaseUrl}?eid=${encodeURIComponent(eid)}`;
    const response = await axios.get(url, {
      headers: CONFIG.headers,
      // 处理压缩响应
      decompress: true
    });

    const { code, data, message } = response.data;

    if (code !== 0) {
      throw new Error(`API返回错误: ${message || '未知错误'}`);
    }

    return data || '无UID数据';
  } catch (error) {
    console.warn(`EID [${eid}] 获取UID失败:`, error.message);
    return `获取失败: ${error.message}`;
  }
}

/**
 * 分批并发处理客户数据
 * @param {Array<{name: string, eid: string}>} customerList - 客户数据列表
 * @returns {Promise<Array<{name: string, eid: string, uid: string}>>} 包含完整信息的数组
 */
async function batchProcessCustomers(customerList) {
  const result = [];
  const totalBatches = Math.ceil(customerList.length / CONFIG.batchSize);

  for (let batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
    // 计算当前批次的起始和结束索引
    const startIndex = batchIndex * CONFIG.batchSize;
    const endIndex = Math.min((batchIndex + 1) * CONFIG.batchSize, customerList.length);
    const currentBatch = customerList.slice(startIndex, endIndex);

    console.log(`\n正在处理第 ${batchIndex + 1}/${totalBatches} 批数据（${startIndex + 1}-${endIndex} 条）`);

    // 并发处理当前批次的所有客户
    const batchResults = await Promise.all(
      currentBatch.map(async (customer) => {
        const { name, eid } = customer;
        console.log(`正在处理客户: ${name} (EID: ${eid})`);
        const uid = await fetchUidByEid(eid);
        return { name, eid, uid };
      })
    );

    // 将当前批次结果添加到总结果中
    result.push(...batchResults);

    // 非最后一批时，添加批次间延迟
    if (batchIndex < totalBatches - 1) {
      console.log(`第 ${batchIndex + 1} 批处理完成，等待 ${CONFIG.batchDelay}ms 后处理下一批...`);
      await new Promise(resolve => setTimeout(resolve, CONFIG.batchDelay));
    }
  }

  return result;
}

/**
 * 将结果导出为Excel文件
 * @param {Array<{name: string, eid: string, uid: string}>} data - 导出数据
 */
function exportToExcel(data) {
  try {
    // 确保输出目录存在
    if (!fs.existsSync(CONFIG.outputDir)) {
      fs.mkdirSync(CONFIG.outputDir, { recursive: true });
    }

    // 创建工作表（指定列顺序）
    const worksheet = XLSX.utils.json_to_sheet(data, {
      header: ['name', 'eid', 'uid'] // 定义列顺序
    });

    // 重命名表头（可选，让表格更易读）
    XLSX.utils.sheet_add_aoa(worksheet, [['客户姓名', '客户EID', '客户UID']], { origin: 'A1' });

    // 创建工作簿并添加工作表
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '客户EID_UID对应表');

    // 输出文件路径
    const outputPath = path.join(CONFIG.outputDir, CONFIG.outputFileName);
    // 写入文件
    XLSX.writeFile(workbook, outputPath);

    console.log(`\nExcel文件已成功导出至: ${outputPath}`);
  } catch (error) {
    console.error('导出Excel失败:', error.message);
    process.exit(1);
  }
}

/**
 * 主函数
 */
async function main() {
  try {
    console.log('开始执行客户EID转UID任务...');

    // 1. 读取客户数据（包含name和eid）
    const customerList = readCustomerDataFromExcel();

    // 2. 分批并发获取UID
    const resultData = await batchProcessCustomers(customerList);

    // 3. 导出Excel
    exportToExcel(resultData);

    console.log('\n任务执行完成！');
  } catch (error) {
    console.error('任务执行失败:', error.message);
    process.exit(1);
  }
}

// 启动脚本
main();