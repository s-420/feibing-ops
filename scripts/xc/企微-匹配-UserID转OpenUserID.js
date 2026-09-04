/**
 * @对象    企微
 * @动作    匹配
 * @风险    低
 * @来源    store-data-extractor/HuoQuUserIDDuiYinOpenUserID.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const fsExtra = require('fs-extra');
const axios = require('axios');

// 配置项
const config = {
  apiUrl: 'https://vinci-api.feibing.tech/sc/v1/sellers/wwa9c5a585540b115b/staffs/convert3rdId',
  headers: {
    'Accept': '*/*',
    'Accept-Encoding': 'gzip, deflate, br',
    'Authorization': 'bearer __VINCI_TOKEN__',
    'Connection': 'keep-alive',
    'Content-Type': 'application/json',
    'User-Agent': 'PostmanRuntime-ApipostRuntime/1.1.0'
  },
  inputFile: path.join(__dirname, '../input', '沪上全部门店对应桌台UserId.xlsx'),
  outputFile: path.join(__dirname, '../output', '桌台对应客服_转换结果.xlsx'),
  userIdColumnName: 'userId', // 输入表格中的userId列名
  resultColumnName: '接口返回结果', // 输出表格中的结果列名
  concurrency: 5, // 并发数量
  maxRetries: 2, // 最大重试次数
  tempDir: path.join(__dirname, '../temp') // 临时文件目录
};

// 确保目录存在
async function ensureDirs() {
  await fsExtra.ensureDir(path.dirname(config.inputFile));
  await fsExtra.ensureDir(path.dirname(config.outputFile));
  await fsExtra.emptyDir(config.tempDir);
  await fsExtra.ensureDir(config.tempDir);
}

// 记录错误日志
function logError(message) {
  const timestamp = new Date().toISOString();
  const errorLogPath = path.join(path.dirname(config.outputFile), 'error_log.txt');
  fs.appendFileSync(errorLogPath, `[${timestamp}] ${message}\n`);
  console.error(`❌ 错误：${message}`);
}

// 读取输入表格中的userId（保留原始顺序）
function readUserIdsFromExcel() {
  try {
    if (!fs.existsSync(config.inputFile)) {
      throw new Error(`输入文件不存在：${config.inputFile}`);
    }

    const workbook = XLSX.readFile(config.inputFile);
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    // 用header: 1确保读取时保留原始列顺序，且不丢失空行（如需）
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    const [headerRow, ...dataRows] = jsonData;

    // 检查是否存在userId列（匹配列名，不区分大小写）
    const userIdColIndex = headerRow.findIndex(
      col => col?.toString().trim().toLowerCase() === config.userIdColumnName.toLowerCase()
    );
    if (userIdColIndex === -1) {
      throw new Error(`输入表格中不存在"${config.userIdColumnName}"列（列名不区分大小写）`);
    }

    // 转换为键值对格式（保留原始顺序）
    const formattedData = dataRows.map(row => {
      const rowObj = {};
      headerRow.forEach((colName, index) => {
        rowObj[colName] = row[index] || ''; // 空值处理为空白字符串
      });
      return rowObj;
    }).filter(row => row[config.userIdColumnName]?.toString().trim()); // 过滤userId为空的行

    console.log(`📋 成功读取${formattedData.length}条有效数据（总行数：${dataRows.length}）`);
    return {
      data: formattedData,
      workbook,
      sheetName: firstSheetName,
      header: headerRow // 保留原始表头，用于生成最终表格
    };
  } catch (error) {
    logError(`读取用户ID失败：${error.message}`);
    throw error;
  }
}

// 调用接口转换userId
async function convertUserId(userId) {
  for (let retry = 1; retry <= config.maxRetries; retry++) {
    try {
      const response = await axios.post(
        `${config.apiUrl}?userId=${encodeURIComponent(userId)}`,
        {}, // POST请求体为空
        {
          headers: config.headers,
          timeout: 10000
        }
      );
      
      return {
        success: true,
        data: response.data
      };
    } catch (error) {
      const errorMsg = error.response?.data?.message || error.message;
      if (retry < config.maxRetries) {
        // 重试前等待时间递增（1s、2s...），避免频繁重试
        await new Promise(resolve => setTimeout(resolve, 1000 * retry));
      } else {
        return {
          success: false,
          error: `调用失败：${errorMsg}`
        };
      }
    }
  }
}

// 处理接口返回结果
function processResult(originalUserId, apiResponse) {
  if (!apiResponse.success) {
    return apiResponse.error;
  }

  const result = apiResponse.data;
  
  // 接口返回错误
  if (result.code !== 0) {
    return `接口错误[${result.code}]：${result.message || '未知错误'}`;
  }
  
  // 结果判断：只要与原始userId不一致就返回结果，一致则标记为“不匹配”
  return result.data === originalUserId ? '不匹配' : (result.data || '无返回数据');
}

// 保存临时结果（按索引保存，确保顺序）
function saveTempResult(index, data) {
  const filePath = path.join(config.tempDir, `result_${index}.json`);
  // 用JSON.stringify确保特殊字符被正确处理
  fs.writeFileSync(filePath, JSON.stringify(data, null, 2), 'utf8');
}

// 读取临时结果（严格按原始索引顺序，不丢失数据）
function readTempResults(totalCount) {
  const results = [];
  for (let i = 0; i < totalCount; i++) {
    const filePath = path.join(config.tempDir, `result_${i}.json`);
    if (fs.existsSync(filePath)) {
      try {
        // 读取并解析临时JSON文件
        const tempData = JSON.parse(fs.readFileSync(filePath, 'utf8'));
        results.push(tempData);
      } catch (e) {
        // 解析失败时，保留原始行数据，标记错误
        results.push({ ...results[i], [config.resultColumnName]: `临时文件解析失败：${e.message}` });
      }
    } else {
      // 无临时文件时，标记为“未处理”
      results.push({ [config.resultColumnName]: '未处理' });
    }
  }
  return results;
}

// 处理单个userId（按索引关联，确保顺序）
async function processSingleUserId(index, row) {
  try {
    const originalUserId = row[config.userIdColumnName]?.toString().trim() || '空';
    
    console.log(`处理第${index + 1}条数据，userId: ${originalUserId}`);
    
    // 调用转换接口
    const apiResponse = await convertUserId(originalUserId);
    
    // 处理结果
    const result = processResult(originalUserId, apiResponse);
    // 保留原始行的所有字段，新增结果列
    const resultRow = { ...row, [config.resultColumnName]: result };
    
    // 按索引保存临时结果（关键：索引与原始数据一致）
    saveTempResult(index, resultRow);
    
    console.log(`完成第${index + 1}条数据，结果：${result.slice(0, 30)}${result.length > 30 ? '...' : ''}`);
  } catch (error) {
    const errorMsg = `处理失败：${error.message}`;
    logError(`第${index + 1}条数据${errorMsg}`);
    // 异常时仍保存结果，确保索引不缺失
    const resultRow = { ...row, [config.resultColumnName]: errorMsg };
    saveTempResult(index, resultRow);
  }
}

// 分批并发处理（原生Promise实现，不依赖外部库）
async function processInBatches(data) {
  const batchSize = config.concurrency;
  const totalBatches = Math.ceil(data.length / batchSize);
  
  console.log(`📦 共${data.length}条数据，分${totalBatches}批处理（每批${batchSize}条，并发数${config.concurrency}）`);
  
  for (let batchIdx = 0; batchIdx < totalBatches; batchIdx++) {
    const start = batchIdx * batchSize;
    const end = Math.min(start + batchSize, data.length);
    const currentBatch = data.slice(start, end);
    
    console.log(`\n🚀 开始处理第${batchIdx + 1}/${totalBatches}批（${start + 1}-${end}条）`);
    
    // 批量创建并发任务：每个任务绑定原始索引（start + i），确保临时结果索引正确
    const batchPromises = currentBatch.map((row, i) => 
      processSingleUserId(start + i, row)
    );
    
    // 等待当前批次所有任务完成，再处理下一批（控制并发量）
    await Promise.all(batchPromises);
    
    console.log(`✅ 第${batchIdx + 1}/${totalBatches}批处理完成`);
  }
}

// 优化：合并所有临时JSON数据并生成表格（严格保持原始顺序）
function generateOutputFile(originalWorkbook, sheetName, data, header) {
  try {
    const totalCount = data.length;
    // 1. 读取所有临时结果（按原始索引顺序）
    const tempResults = readTempResults(totalCount);
    
    // 2. 合并：临时结果优先，无临时结果时用原始数据（避免丢失）
    const finalData = tempResults.map((tempRow, index) => {
      return tempRow ? tempRow : { ...data[index], [config.resultColumnName]: '未处理' };
    });
    
    // 3. 生成Excel表格（用原始表头确保列顺序一致）
    // 第一步：创建表头行（原始表头 + 结果列）
    const finalHeader = [...header, config.resultColumnName];
    // 第二步：将finalData转换为“表头-数据”格式（确保列顺序与表头一致）
    const excelData = [
      finalHeader, // 表头行
      ...finalData.map(row => finalHeader.map(col => row[col] || '')) // 数据行（按表头顺序填充）
    ];
    
    // 4. 创建工作表并写入文件
    const newWorksheet = XLSX.utils.aoa_to_sheet(excelData); // 用aoa_to_sheet确保顺序
    originalWorkbook.Sheets[sheetName] = newWorksheet;
    XLSX.writeFile(originalWorkbook, config.outputFile);
    
    console.log(`🎉 合并完成！共${finalData.length}条数据，结果文件：${config.outputFile}`);
    console.log(`📊 结果列说明：${config.resultColumnName}`);
  } catch (error) {
    logError(`生成表格失败：${error.message}`);
    throw error;
  }
}

// 主函数
async function main() {
  try {
    console.log(`🚀 启动userId转换任务（并发数：${config.concurrency}）`);
    
    // 1. 初始化目录（清空临时文件夹，确保输出目录存在）
    await ensureDirs();
    
    // 2. 读取Excel数据（保留原始表头和顺序）
    const { data, workbook, sheetName, header } = readUserIdsFromExcel();
    if (data.length === 0) {
      console.log(`ℹ️ 无有效数据，程序退出`);
      return;
    }
    
    // 3. 分批并发处理所有数据
    await processInBatches(data);
    
    // 4. 合并所有临时JSON结果，生成最终表格
    generateOutputFile(workbook, sheetName, data, header);
    
    console.log(`\n✅ 所有任务完成！`);
  } catch (error) {
    console.error(`💥 任务中断：${error.message}`);
    process.exit(1);
  } finally {
    console.log(`\n📝 任务结束（临时文件路径：${config.tempDir}）`);
  }
}

// 启动程序
main().catch(error => {
  logError(`主函数异常：${error.message}`);
  process.exit(1);
});