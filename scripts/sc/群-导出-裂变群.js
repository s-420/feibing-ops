/**
 * @对象    群
 * @动作    导出
 * @风险    低
 * @来源    store-data-extractor/DaoChuQuanBuLieBianQun.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const fsExtra = require('fs-extra');
const axios = require('axios');

// 导出飞冰全部裂变群汇总

// 配置
const config = {
  outputDir: path.join(__dirname, '../output'), // 最终汇总表目录
  tempDir: path.join(__dirname, '../temp_excels'), // 临时分表目录
  finalFile: `全部裂变群汇总_${new Date().toISOString().replace(/[:.]/g, '-').slice(0,19)}.xlsx`, // 最终汇总表名称
  sheetName: '数据列表',
  apiUrl: 'https://vinci-api.feibing.tech/sc/v1/sellers/wpV3cNDAAA55j60tcJ8ds1N8BHXD3D_g/configs',
  headers: {
    'accept': '*/*',
    'accept-language': 'zh-CN,zh;q=0.9',
    "authorization": "bearer __VINCI_TOKEN__",
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
  cookies: {
    '_clck': 'sj51u%7C2%7Cfxn%7C0%7C1974',
    'x-token': '__VINCI_TOKEN__'
  },
  pageSize: 100, // 每页100条
  delay: 2000, // 20秒请求一次
  total: 433
};

// 确保输出目录和临时目录存在
async function ensureDirs() {
  await fsExtra.ensureDir(config.outputDir);
  await fsExtra.ensureDir(config.tempDir);
}

// 转换cookie为字符串
function cookiesToString(cookies) {
  return Object.entries(cookies).map(([key, value]) => `${key}=${value}`).join('; ');
}

// 发送API请求
async function fetchData(pageNum) {
  try {
    console.log(`📡 正在请求第${pageNum}页数据...`);
    const response = await axios.get(config.apiUrl, {
      params: {
        pageNum,
        pageSize: config.pageSize,
        type: 'GROUP_CHAT_JOIN_WAY',
        sort: 'createdTime,desc'
      },
      headers: {
        ...config.headers,
        'Cookie': cookiesToString(config.cookies)
      }
    });
    
    const dataCount = response.data?.data?.length || 0;
    console.log(`📥 第${pageNum}页请求成功，返回${dataCount}条数据`);
    return response.data;
  } catch (error) {
    console.error(`❌ 第${pageNum}页请求失败:`, error.message);
    throw error;
  }
}

// 处理数据
function processData(items, pageNum) {
  console.log(`开始处理第${pageNum}页的${items.length}条数据`);
  
  return items.map((item, index) => {
    let valueData = {};
    try {
      valueData = JSON.parse(item.value || '{}');
     
    } catch (e) {
      console.error(`❌ 第${pageNum}页第${index+1}条数据解析失败:`, e.message);
      valueData = {};
    }
    
    const joinWay = valueData.join_way || {};
    const chatIdList = Array.isArray(joinWay.chat_id_list) ? joinWay.chat_id_list : [];
    
    configRes = JSON.parse(item.metadata.configRes || '{}').join_way.remark;
    configSource = JSON.parse(item.metadata.configSource || '{}').remark;
    
    return {
      'key': item.key || '无key',
      'room_base_name': joinWay.room_base_name || '无名称',
      'shop_id':item.metadata.shopId ||'无门店',
      'remark_shop_name':joinWay.remark || configRes || configSource || '无门店名称',
      'chat_id_list': chatIdList.join(',') || '无ID',
      'chat_id_count': chatIdList.length,
      'link': item.metadata.qrCode || '无链接',
      'qr_code': joinWay.qr_code || '无二维码',
    };
  });
}

// 生成单页Excel（保存到临时目录）
function generatePageExcel(data, pageNum) {
  if (data.length === 0) {
    console.warn(`⚠️ 第${pageNum}页无数据，不生成Excel`);
    return null;
  }
  
  // 临时文件路径：temp_excels/第1页.xlsx
  const tempFilePath = path.join(config.tempDir, `第${pageNum}页.xlsx`);
  
  try {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(data, {
      header: ['key', 'room_base_name', 'chat_id_list', 'chat_id_count', 'qr_code', 'id', 'createdTime']
    });
    XLSX.utils.book_append_sheet(workbook, worksheet, config.sheetName);
    XLSX.writeFile(workbook, tempFilePath);
    
    console.log(`✅ 第${pageNum}页数据已保存到临时文件: ${tempFilePath}`);
    return tempFilePath;
  } catch (error) {
    console.error(`❌ 生成第${pageNum}页Excel失败:`, error.message);
    throw error;
  }
}

// 汇总所有临时Excel到一个文件
async function mergeAllExcels() {
  console.log('\n开始汇总所有临时Excel文件...');
  
  // 获取临时目录下的所有Excel文件
  const tempFiles = fs.readdirSync(config.tempDir)
    .filter(file => file.endsWith('.xlsx') && file.startsWith('第') && file.includes('页'))
    .sort((a, b) => {
      // 按页码排序
      const pageA = parseInt(a.match(/第(\d+)页/)[1]);
      const pageB = parseInt(b.match(/第(\d+)页/)[1]);
      return pageA - pageB;
    });
  
  if (tempFiles.length === 0) {
    console.warn('⚠️ 未找到任何临时Excel文件，无法汇总');
    return null;
  }
  
  console.log(`找到${tempFiles.length}个临时Excel文件，开始合并...`);
  
  // 汇总所有数据
  let allData = [];
  tempFiles.forEach(file => {
    const filePath = path.join(config.tempDir, file);
    const workbook = XLSX.readFile(filePath);
    const worksheet = workbook.Sheets[config.sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet);
    allData = allData.concat(data);
    console.log(`合并${file}，新增${data.length}条数据`);
  });
  
  // 生成最终汇总文件
  const finalFilePath = path.join(config.outputDir, config.finalFile);
  const finalWorkbook = XLSX.utils.book_new();
  const finalWorksheet = XLSX.utils.json_to_sheet(allData, {
    header: ['key', 'room_base_name', 'chat_id_list', 'chat_id_count', 'qr_code', 'id', 'createdTime']
  });
  XLSX.utils.book_append_sheet(finalWorkbook, finalWorksheet, config.sheetName);
  XLSX.writeFile(finalWorkbook, finalFilePath);
  
  console.log(`🎉 所有数据汇总完成，共${allData.length}条，保存到: ${finalFilePath}`);
  return finalFilePath;
}

// 延迟函数
function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// 主函数
async function main() {
  try {
    // 初始化目录
    await ensureDirs();
    
    // 计算总页数
    const totalPages = Math.ceil(config.total / config.pageSize);
    console.log(`🚀 开始执行数据获取，共需获取${totalPages}页，总计${config.total}条数据`);
    
    // 逐页获取并生成临时Excel
    for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
      try {
        const apiResponse = await fetchData(pageNum);
        
        if (apiResponse.code !== 0) {
          console.error(`⚠️ 第${pageNum}页API返回错误: ${apiResponse.message}`);
          await delay(config.delay);
          continue;
        }
        
        const rawData = apiResponse.data || [];
        if (rawData.length === 0) {
          console.warn(`⚠️ 第${pageNum}页无数据，跳过`);
          await delay(config.delay);
          continue;
        }
        
        // 处理并生成临时Excel
        const processedData = processData(rawData, pageNum);
        await generatePageExcel(processedData, pageNum);
        
        // 最后一页不需要延迟
        if (pageNum < totalPages) {
          console.log(`⏳ 等待${config.delay/1000}秒后继续下一页...\n`);
          await delay(config.delay);
        }
      } catch (pageError) {
        console.error(`⚠️ 第${pageNum}页处理出错，重试...`);
        await delay(config.delay);
      }
    }
    
    // 汇总所有临时文件
    await mergeAllExcels();
    
    console.log('\n🎉 所有操作完成！');
  } catch (error) {
    console.error('\n💥 执行中断:', error.message);
  }
}

// 启动程序
main();