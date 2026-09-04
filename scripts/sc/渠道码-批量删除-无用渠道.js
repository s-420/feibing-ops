/**
 * @对象    渠道码
 * @动作    批量写
 * @风险    高
 * @来源    store-data-extractor/deleteWuYongQuDao.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const axios = require('axios');

// 配置项 - 新增批次大小配置
const config = {
  excelPath: path.join(__dirname, '../input/删除码.xlsx'),
  idColumnName: '渠道码ID',
  maxRows: 0, // 0表示全部执行
  batchSize: 10, // 每批并发数量
  baseUrl: 'https://scrm.feibing.tech/wework-scrm/contact',
  headers: {
    'accept': '*/*',
    'accept-language': 'zh-CN,zh;q=0.9',
    'origin': 'https://scrm.feibing.tech',
    'priority': 'u=1, i',
    'referer': 'https://scrm.feibing.tech/',
    'sec-ch-ua': '"Not;A=Brand";v="99", "Google Chrome";v="139", "Chromium";v="139"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"macOS"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36',
    'cookie': '_clck=sj51u%5E2%5Efyn%5E0%5E1974; _bl_uid=31b912ca-5464-411f-a3bf-c6bb611f4038; x-token=__VINCI_TOKEN__; acw_tc=0bca30f617588714384675639ea0e11d2d725d103bca6972d974a25c3b3e97'
  },
  delay: 1000, // 批次之间的延迟时间(毫秒)
};

const errorRecords = [];

// 读取Excel中的渠道码ID
function readChannelIdsFromExcel() {
  try {
    if (!fs.existsSync(config.excelPath)) {
      throw new Error(`文件不存在: ${config.excelPath}`);
    }

    const workbook = XLSX.readFile(config.excelPath);
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    if (!jsonData || jsonData.length === 0) {
      throw new Error('Excel文件中没有数据');
    }

    const firstRow = jsonData[0];
    const idField = Object.keys(firstRow).find(key => 
      key.trim() === config.idColumnName.trim()
    );

    if (!idField) {
      throw new Error(`未找到名为"${config.idColumnName}"的列`);
    }

    const channelIds = jsonData.map((row, index) => ({
      id: row[idField],
      rowNumber: index + 2
    }));

    console.log(`成功读取到${channelIds.length}个渠道码ID`);
    return channelIds;
  } catch (error) {
    console.error('读取Excel失败:', error.message);
    process.exit(1);
  }
}

// 调用删除接口
async function callDeleteApi(id, rowNumber) {
  try {
    console.log(`正在处理第${rowNumber}行: ${id}`);
    const url = `${config.baseUrl}?id=${id}`;
    
    const response = await axios.delete(url, {
      headers: config.headers,
      withCredentials: true
    });
    if (response.status >= 200 && response.status < 300) {
      console.log(`第${rowNumber}行删除成功`);
      return { success: true, rowNumber, id };
    } else {
      throw new Error(`状态码异常: ${response.status}`);
    }
  } catch (error) {
    const errorMsg = error.response 
      ? `${error.response.status} - ${JSON.stringify(error.response.data)}`
      : error.message;
      
    console.error(`第${rowNumber}行失败: ${errorMsg}`);
    errorRecords.push({
      rowNumber,
      channelId: id,
      error: error.message,
      status: error.response?.status || '未知'
    });
    return { success: false, rowNumber, id, error: errorMsg };
  }
}

// 批量执行（分批并发处理）
async function batchDelete() {
  try {
    const channelIds = readChannelIdsFromExcel();
    const totalCount = config.maxRows > 0 
      ? Math.min(config.maxRows, channelIds.length)
      : channelIds.length;

    // 截取需要处理的部分
    const processData = channelIds.slice(0, totalCount);
    // 计算总批次
    const totalBatches = Math.ceil(processData.length / config.batchSize);
    
    console.log(`开始处理${processData.length}条记录，共${totalBatches}批，每批${config.batchSize}条`);
    
    // 按批次处理
    for (let batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
      // 计算当前批次的起始和结束索引
      const startIndex = batchIndex * config.batchSize;
      const endIndex = Math.min(startIndex + config.batchSize, processData.length);
      const currentBatch = processData.slice(startIndex, endIndex);
      
      console.log(`\n===== 开始处理第${batchIndex + 1}/${totalBatches}批 (${startIndex + 1}-${endIndex}条) =====`);
      
      // 并发执行当前批次的所有请求
      const batchPromises = currentBatch.map(item => 
        callDeleteApi(item.id, item.rowNumber)
      );
      
      // 等待当前批次所有请求完成（无论成功失败）
      await Promise.allSettled(batchPromises);
      
      console.log(`第${batchIndex + 1}/${totalBatches}批处理完成`);
      
      // 不是最后一批则添加批次间延迟
      if (batchIndex < totalBatches - 1) {
        console.log(`等待${config.delay}ms后处理下一批...`);
        await new Promise(resolve => setTimeout(resolve, config.delay));
      }
    }
    
    // 输出结果
    console.log('\n===== 执行结果汇总 =====');
    console.log(`总处理: ${processData.length}条`);
    console.log(`成功: ${processData.length - errorRecords.length}条`);
    console.log(`失败: ${errorRecords.length}条`);
    
    if (errorRecords.length > 0) {
      console.log('\n===== 错误详情 =====');
      errorRecords.forEach(record => {
        console.log(`行号: ${record.rowNumber}, ID: ${record.channelId}, 状态: ${record.status}`);
      });
      
      const errorLogPath = path.join(__dirname, 'delete_errors.log');
      fs.writeFileSync(errorLogPath, JSON.stringify(errorRecords, null, 2));
      console.log(`\n错误记录已保存到: ${errorLogPath}`);
    }
    
    console.log('\n所有操作完成');
  } catch (error) {
    console.error('批量操作失败:', error.message);
  }
}

batchDelete();