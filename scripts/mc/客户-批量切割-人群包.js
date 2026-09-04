/**
 * @对象    客户
 * @动作    批量写
 * @风险    高
 * @来源    store-data-extractor/QieGeRenQunBao.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const fs = require('fs').promises;
const fsSync = require('fs');
const path = require('path');
const readline = require('readline');

// 配置参数
const inputDir = path.join(__dirname, '../input'); // 输入文件夹路径
const outputDir = path.join(__dirname, '../output'); // 输出文件夹路径
const MAX_RECORDS_PER_FILE = 100000; // 每个文件最大记录数

/**
 * 检查文件是否为CSV格式
 * @param {string} filename 文件名
 * @returns {boolean} 是否为CSV文件
 */
function isCsvFile(filename) {
  return path.extname(filename).toLowerCase() === '.csv';
}

/**
 * 获取输出文件路径，保持与输入目录相同的相对结构
 * @param {string} csvPath CSV文件路径
 * @returns {string} 输出目录路径
 */
function getOutputDir(csvPath) {
  // 获取文件相对于input目录的路径
  const relativePath = path.relative(inputDir, path.dirname(csvPath));
  // 构建output目录下的对应路径
  return path.join(outputDir, relativePath);
}

/**
 * 处理单个CSV文件：删除表头并按10万条记录分割，输出到output文件夹
 * @param {string} csvPath CSV文件路径
 */
async function processCsvFile(csvPath) {
  try {
    const fileName = path.basename(csvPath, '.csv');
    // 获取输出目录并确保目录存在
    const outputDirPath = getOutputDir(csvPath);
    await fs.mkdir(outputDirPath, { recursive: true });
    
    // 创建readline接口用于逐行读取（适合大文件处理）
    const rl = readline.createInterface({
      input: fsSync.createReadStream(csvPath),
      crlfDelay: Infinity
    });
    
    let lineNumber = 0;
    let currentFileNumber = 1;
    let currentWriteStream = null;
    let recordCount = 0;
    
    // 创建第一个输出文件
    const firstOutputPath = path.join(outputDirPath, `${fileName}_${currentFileNumber}.txt`);
    currentWriteStream = fsSync.createWriteStream(firstOutputPath);
    console.log(`创建文件: ${path.relative(outputDir, firstOutputPath)}`);
    
    for await (const line of rl) {
      lineNumber++;
      
      // 跳过表头（第一行）
      if (lineNumber === 1) {
        continue;
      }
      
      // 如果当前文件记录数达到上限，创建新文件
      if (recordCount >= MAX_RECORDS_PER_FILE) {
        currentWriteStream.end();
        currentFileNumber++;
        recordCount = 0;
        
        const newOutputPath = path.join(outputDirPath, `${fileName}_${currentFileNumber}.txt`);
        currentWriteStream = fsSync.createWriteStream(newOutputPath);
        console.log(`创建文件: ${path.relative(outputDir, newOutputPath)}`);
      }
      
      // 写入记录
      currentWriteStream.write(line + '\n');
      recordCount++;
    }
    
    // 关闭最后一个文件流
    if (currentWriteStream) {
      currentWriteStream.end();
    }
    
    const totalRecords = recordCount + (currentFileNumber - 1) * MAX_RECORDS_PER_FILE;
    console.log(`处理完成: ${path.basename(csvPath)} -> 共生成 ${currentFileNumber} 个文件，总记录数 ${totalRecords}`);
    return true;
  } catch (error) {
    console.error(`处理文件 ${path.basename(csvPath)} 时出错:`, error.message);
    return false;
  }
}

/**
 * 递归处理目录下的所有CSV文件
 * @param {string} dir 目录路径
 */
async function processDirectory(dir) {
  try {
    // 检查目录是否存在
    await fs.access(dir);
  } catch {
    console.error(`错误: 目录不存在 - ${dir}`);
    return;
  }

  try {
    // 确保输出目录存在
    await fs.mkdir(outputDir, { recursive: true });
    
    // 读取目录内容
    const entries = await fs.readdir(dir, { withFileTypes: true });
    
    for (const entry of entries) {
      const fullPath = path.join(dir, entry.name);
      
      if (entry.isDirectory()) {
        // 递归处理子目录
        await processDirectory(fullPath);
      } else if (entry.isFile() && isCsvFile(entry.name)) {
        // 处理CSV文件
        await processCsvFile(fullPath);
      }
    }
    
    console.log('所有文件处理完成');
  } catch (error) {
    console.error('处理目录时出错:', error.message);
  }
}

// 主函数
async function main() {
  console.log('开始处理CSV文件...');
  console.log(`输入目录: ${inputDir}`);
  console.log(`输出目录: ${outputDir}`);
  console.log(`每个文件最大记录数: ${MAX_RECORDS_PER_FILE}`);
  await processDirectory(inputDir);
}

// 执行主函数
main();
