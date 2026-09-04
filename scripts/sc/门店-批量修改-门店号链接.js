/**
 * @对象    门店
 * @动作    批量写
 * @风险    高
 * @来源    store-data-extractor/XiuGaiWeiXinMenDianHaoLink.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const axios = require('axios');

// 配置参数
const config = {
  excelPath: path.join(__dirname, '../input', '门店ID和入群链接.xlsx'),
  logPath: path.join(__dirname, '../output', '执行日志.txt'), // 日志文件路径
  baseUrl: 'https://vinci-api.feibing.tech/sc/v1/sellers/wwa9c5a585540b115b/configs',
  // 需要被替换的目标链接
  targetLink: 'https://work.weixin.qq.com/gm/c1746424b5a8018b0afa156a12e62a5e',
  // https://work.weixin.qq.com/gm/4b5c9e69ff72e3abfd40cf8a1c5019eb   茶瀑布
  // 'https://work.weixin.qq.com/gm/c1746424b5a8018b0afa156a12e62a5e', 轻享
  // 请求头信息
  headers: {
    'accept': '*/*',
    'accept-language': 'zh-CN,zh;q=0.9',
    'authorization': 'bearer __VINCI_TOKEN__',
    'origin': 'https://connect.feibing.tech',
    'priority': 'u=1, i',
    'referer': 'https://connect.feibing.tech/',
    'sec-ch-ua': '"Not;A=Brand";v="99", "Google Chrome";v="139", "Chromium";v="139"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"macOS"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-site',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36',
    'cookie': '_clck=sj51u%5E2%5Efyn%5E0%5E1974; x-token=__VINCI_TOKEN__'
  }
};

// 读取Excel文件
function readExcelFile(filePath) {
  try {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    return xlsx.utils.sheet_to_json(worksheet);
  } catch (error) {
    const errorMsg = `读取Excel文件失败: ${error.message}`;
    logToFile(errorMsg);
    console.error(errorMsg); // 只在控制台显示关键错误
    throw error;
  }
}

// 记录日志到TXT文件
function logToFile(message) {
  try {
    // 确保output目录存在
    if (!fs.existsSync(path.dirname(config.logPath))) {
      fs.mkdirSync(path.dirname(config.logPath), { recursive: true });
    }
    
    // 获取当前时间并格式化
    const timestamp = new Date().toLocaleString('zh-CN', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false
    });
    
    // 格式化日志内容
    const logEntry = `[${timestamp}] ${message}\n`;
    
    // 追加写入日志
    fs.appendFileSync(config.logPath, logEntry, 'utf8');
  } catch (error) {
    console.error('写入日志失败:', error.message); // 日志写入失败需要在控制台显示
  }
}

// 获取配置详情
async function getConfigDetails(storeId) {
  try {
    const url = `${config.baseUrl}?type=MINI_APP_STYLE_TYPE&key=${storeId}_pre_h5_config_wx`;
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
    const errorMsg = `获取门店 ${storeId} 配置失败: ${error.message}`;
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
    const errorMsg = `更新配置 ${configId} 失败: ${error.response?.data || error.message}`;
    logToFile(errorMsg);
    throw error;
  }
}

// 替换链接并准备更新数据 - 只替换特定目标链接
function prepareUpdateData(originalConfig, newLink, storeId) {
  try {
    const valueObj = JSON.parse(originalConfig.value);
    let replaceCount = 0;
    let originalLinks = []; // 存储原始链接
    
    // 1. 替换json.list中第四个元素的path（只替换目标链接）
    if (valueObj.json && valueObj.json.list && valueObj.json.list.length >= 4) {
      const targetItem = valueObj.json.list[3];
      if (targetItem.children && targetItem.children[0]) {
        if (targetItem.children[0].path === config.targetLink) {
          originalLinks.push(targetItem.children[0].path);
          logToFile(`门店 ${storeId} - 替换JSON中的链接: ${targetItem.children[0].path} -> ${newLink}`);
          targetItem.children[0].path = newLink;
          replaceCount++;
        } else {
          logToFile(`门店 ${storeId} - JSON中的链接不是目标链接，无需替换: ${targetItem.children[0].path}`);
        }
      } else {
        const errorMsg = '未找到需要替换的链接位置（JSON部分）';
        logToFile(errorMsg);
        throw new Error(errorMsg);
      }
    } else {
      const errorMsg = '配置结构不符合预期（JSON部分）';
      logToFile(errorMsg);
      throw new Error(errorMsg);
    }
    
    // 2. 替换html中的链接（只替换目标链接）
    if (valueObj.html) {
      const oldLinkRegex = new RegExp(config.targetLink, 'g');
      const matches = valueObj.html.match(oldLinkRegex);
      
      if (matches && matches.length > 0) {
        originalLinks = [...originalLinks, ...matches];
        logToFile(`门店 ${storeId} - 替换HTML中的链接，共${matches.length}处: ${config.targetLink} -> ${newLink}`);
        valueObj.html = valueObj.html.replace(oldLinkRegex, newLink);
        replaceCount += matches.length;
      } else {
        logToFile(`门店 ${storeId} - HTML中未找到目标链接，无需替换`);
      }
    }
    
    if (replaceCount === 0) {
      logToFile(`门店 ${storeId} - 未找到任何需要替换的目标链接，不执行更新`);
      return { 
        updateData: null,
        originalLinks: []
      };
    }
    
    const updatedValue = JSON.stringify(valueObj);
    
    return {
      updateData: {
        type: originalConfig.type,
        key: originalConfig.key,
        value: updatedValue
      },
      originalLinks: originalLinks
    };
  } catch (error) {
    const errorMsg = `处理配置数据失败: ${error.message}`;
    logToFile(errorMsg);
    throw error;
  }
}

// 处理单个门店
async function processStore(store) {
  try {
    const storeId = store['店铺ID'];
    const newLink = store['入群链接'];
    logToFile(`\n开始处理门店: ${storeId}`);
    console.log(`处理中: ${storeId}`); // 控制台只显示正在处理的门店ID
    
    // 获取配置详情
    const configDetails = await getConfigDetails(storeId);
    
    // 准备更新数据
    const { updateData, originalLinks } = prepareUpdateData(configDetails, newLink, storeId);
    
    // 如果无需更新
    if (!updateData) {
      logToFile(`门店 ${storeId} - 处理完成，未进行更新`);
      console.log(`已完成: ${storeId} (未更新)`);
      return;
    }
    
    // 执行更新
    const updateResult = await updateConfig(configDetails.id, updateData);
    
    if (updateResult.code === 0) {
      // 记录详细的替换信息
      logToFile(`门店 ${storeId} - 更新成功`);
      logToFile(`门店 ${storeId} - 原始链接: ${originalLinks.join('; ')}`);
      logToFile(`门店 ${storeId} - 替换为: ${newLink}`);
      console.log(`已完成: ${storeId} (成功)`);
    } else {
      logToFile(`门店 ${storeId} - 更新失败: ${updateResult.message || '更新失败，未知原因'}`);
      console.log(`已完成: ${storeId} (失败)`);
    }
    
  } catch (error) {
    logToFile(`门店 ${store['店铺ID']} - 处理失败: ${error.message}`);
    console.log(`已完成: ${store['店铺ID']} (失败)`);
  }
}

// 主函数
async function main() {
  try {
    // 记录程序开始
    logToFile('===== 程序开始执行 =====');
    console.log('程序开始执行...');
    
    // 检查input目录是否存在
    if (!fs.existsSync(path.dirname(config.excelPath))) {
      fs.mkdirSync(path.dirname(config.excelPath), { recursive: true });
      const errorMsg = '已创建input目录，请将Excel文件放入该目录后重新运行';
      logToFile(errorMsg);
      console.error(errorMsg);
      process.exit(1);
    }
    
    // 读取Excel数据
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
        continue;
      }
      
      await processStore(store);
    }
    
    logToFile('===== 程序执行结束 =====\n');
    console.log(`\n程序执行完成，详细日志已保存到: ${config.logPath}`);
    
  } catch (error) {
    const errorMsg = `程序执行出错: ${error.message}`;
    logToFile(errorMsg);
    console.error(errorMsg);
    process.exit(1);
  }
}

// 启动程序
main();
