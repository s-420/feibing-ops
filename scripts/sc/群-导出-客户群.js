/**
 * @对象    群
 * @动作    导出
 * @风险    低
 * @来源    store-data-extractor/DaoChuQUanBuKeHuQun.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const fsExtra = require('fs-extra');
const axios = require('axios');

// 导出飞冰全部客户群脚本（增加关联门店字段及群主ID统计）
const config = {
  outputDir: path.join(__dirname, '../output/客户群汇总表'),
  tempDir: path.join(__dirname, '../output/客户群临时表'),
  finalFile: `全部客户群汇总_${new Date().toISOString().replace(/[:.]/g, '-').slice(0,19)}.xlsx`,
  sheetName: '活动列表',
  apiUrl: 'https://vinci-api.feibing.tech/sc/v2/sellers/wpV3cNDAAA7gXDSpALtCGXgFxSpYMBQQ/activities',
  headers: {
    'accept': '*/*',
    'accept-language': 'zh-CN,zh;q=0.9',
    'origin': 'https://connect.feibing.tech',
    "authorization": "bearer __VINCI_TOKEN__",
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
  pageSize: 100,  // 保持每页100条
  total: 484,
  columns: ['id', 'name','群主ID','门店ID', '关联门店','群人数','创群时间'] // 新增“关联门店”列
};

// 确保输出目录和临时目录存在
async function ensureDirs() {
  await fsExtra.ensureDir(config.outputDir);
  await fsExtra.ensureDir(config.tempDir);
}

// 转换cookie为字符串（与curl格式一致）
function cookiesToString(cookies) {
  return Object.entries(cookies).map(([key, value]) => `${key}=${value}`).join('; ');
}

// 发送API请求（严格匹配curl参数）
async function fetchData(pageNum) {
  try {
    console.log(`📡 正在请求第${pageNum}页活动数据...`);
    const response = await axios.get(config.apiUrl, {
      params: {
        pageNum,  // 页码参数（与curl一致）
        pageSize: config.pageSize,  // 每页条数
        type: 'GROUP_CHAT',  // 固定参数
        sort: 'createdTime,desc',  // 排序参数
        status: 'PUBLISHED',  // 状态参数
        catalogId: '',  // 空目录ID
        ownerId: 'woV3cNDAAArhr9E4nFcLh1X7dje0Anwg'
      },
      headers: {
        ...config.headers,
        'Cookie': cookiesToString(config.cookies)  // 附加cookie
      }
    });
    
    const dataCount = response.data?.data?.length || 0;
    console.log(`📥 第${pageNum}页请求成功，返回${dataCount}条活动数据`);
    return response.data;
  } catch (error) {
    console.error(`❌ 第${pageNum}页请求失败:`, error.message);
    // 打印完整错误信息辅助调试
    if (error.response) {
      console.error(`响应状态码: ${error.response.status}`);
      console.error(`响应数据: ${JSON.stringify(error.response.data)}`);
    }
    throw error;
  }
}

// 处理数据（增加关联门店字段）
function processData(items, pageNum) {
  console.log(`开始处理第${pageNum}页的${items.length}条活动数据`);
  
  return items.map((item, index) => {
    // 提取shops数组中的所有门店名称
    let shopNames = [];
    let shopID = [];
    if (Array.isArray(item.shops)) {
      // 遍历shops数组，收集所有name字段
      shopNames = item.shops.map(shop => shop.name || '未知门店');
      shopID = item.shops.map(shop => shop.id || '未知ID');
    }
    
    // 合并门店名称为字符串（用逗号分隔）
    const relatedShops = shopNames.join('，') || '无关联门店';
    const relatedShopIDs = shopID.join('，') || '无关联门店ID';
    
    return {
      'id': item.id || '无ID',
      'name': item.name || '无名称',
      '群主ID': item.owner?.id || '无群主ID',
      "门店ID": relatedShopIDs, // 门店ID
      '关联门店': relatedShops, // 关联门店
      '群人数': item.realNum || 0,
      '创群时间': item.createdTime ? new Date(item.createdTime).toLocaleString() : '未知时间'
    };
  });
}

// 生成单页Excel（保存到临时目录）
function generatePageExcel(data, pageNum) {
  if (data.length === 0) {
    console.warn(`⚠️ 第${pageNum}页无活动数据，不生成Excel`);
    return null;
  }
  
  const tempFilePath = path.join(config.tempDir, `第${pageNum}页活动数据.xlsx`);
  
  try {
    const workbook = XLSX.utils.book_new();
    // 按配置的列顺序写入（包含新增的“关联门店”）
    const worksheet = XLSX.utils.json_to_sheet(data, {
      header: config.columns
    });
    XLSX.utils.book_append_sheet(workbook, worksheet, config.sheetName);
    XLSX.writeFile(workbook, tempFilePath);
    
    console.log(`✅ 第${pageNum}页活动数据已保存到临时文件: ${tempFilePath}`);
    return tempFilePath;
  } catch (error) {
    console.error(`❌ 生成第${pageNum}页活动Excel失败:`, error.message);
    throw error;
  }
}

// 汇总所有临时Excel到一个文件，并统计群主ID总数
async function mergeAllExcels() {
  console.log('\n开始汇总所有活动数据临时Excel文件...');
  
  const tempFiles = fs.readdirSync(config.tempDir)
    .filter(file => file.endsWith('.xlsx') && file.includes('活动数据'))
    .sort((a, b) => {
      const pageA = parseInt(a.match(/第(\d+)页/)[1]);
      const pageB = parseInt(b.match(/第(\d+)页/)[1]);
      return pageA - pageB;
    });
  
  if (tempFiles.length === 0) {
    console.warn('⚠️ 未找到任何活动数据临时Excel文件，无法汇总');
    return null;
  }
  
  console.log(`找到${tempFiles.length}个活动数据临时Excel文件，开始合并...`);
  
  let allData = [];
  // 用于统计群主ID的集合（自动去重）
  const ownerIds = new Set();
  
  tempFiles.forEach(file => {
    const filePath = path.join(config.tempDir, file);
    const workbook = XLSX.readFile(filePath);
    const worksheet = workbook.Sheets[config.sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet);
    allData = allData.concat(data);
    
    // 收集当前页的群主ID
    data.forEach(item => {
      if (item['群主ID'] && item['群主ID'] !== '无群主ID') {
        ownerIds.add(item['群主ID']);
      }
    });
    
    console.log(`合并${file}，新增${data.length}条活动数据`);
  });
  
  // 添加统计行
  const totalOwnerCount = ownerIds.size;
  allData.push({
    'id': '总计',
    'name': '',
    '群主ID': totalOwnerCount,
    '关联门店': '个不同的群主',
    '群人数': '',
    '创群时间': ''
  });
  
  // 生成最终汇总文件
  const finalFilePath = path.join(config.outputDir, config.finalFile);
  const finalWorkbook = XLSX.utils.book_new();
  const finalWorksheet = XLSX.utils.json_to_sheet(allData, {
    header: config.columns
  });
  XLSX.utils.book_append_sheet(finalWorkbook, finalWorksheet, config.sheetName);
  XLSX.writeFile(finalWorkbook, finalFilePath);
  
  console.log(`🎉 所有活动数据汇总完成，共${allData.length - 1}条记录，不同的群主ID总数为${totalOwnerCount}个，保存到: ${finalFilePath}`);
  return finalFilePath;
}

// 主函数
async function main() {
  try {
    await ensureDirs();
    const totalPages = Math.ceil(config.total / config.pageSize);
    console.log(`🚀 开始执行活动数据获取，共需获取${totalPages}页，总计约${config.total}条数据`);
    
    for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
      try {
        const apiResponse = await fetchData(pageNum);
        
        // 检查API返回状态（根据实际接口调整判断逻辑）
        if (apiResponse.code !== 0) {
          console.error(`⚠️ 第${pageNum}页API返回错误: ${apiResponse.message || '未知错误'}`);
          continue; // 移除延迟，直接继续
        }
        
        const rawData = apiResponse.data || [];
        if (rawData.length === 0) {
          console.warn(`⚠️ 第${pageNum}页无活动数据，跳过`);
          if (pageNum > 1) {
            console.log("推测已获取全部数据，提前结束");
            break;
          }
          continue; // 移除延迟，直接继续
        }
        
        const processedData = processData(rawData, pageNum);
        await generatePageExcel(processedData, pageNum);
        
        // 移除分页之间的延迟
      } catch (pageError) {
        console.error(`⚠️ 第${pageNum}页处理出错，重试...`);
        // 移除错误重试时的延迟
        pageNum--; // 重试当前页
      }
    }
    
    await mergeAllExcels();
    console.log('\n🎉 所有活动数据操作完成！');
  } catch (error) {
    console.error('\n💥 执行中断:', error.message);
  }
}

// 启动程序
main();
    