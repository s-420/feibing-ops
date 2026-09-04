/**
 * @对象    通用
 * @动作    导出
 * @风险    低
 * @来源    store-data-extractor/XiaZaiTuPian.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const fsExtra = require('fs-extra');
const axios = require('axios');

// 配置：新增重试、校验核心参数
const config = {
  tempDir: path.join(__dirname, '../temp_data_two'),
  outputDir: path.join(__dirname, '../output/图片仓库'),
  inputFile: path.join(__dirname, '../input/下载图片链接.xlsx'),
  imageUrlColumnName: 'preQrCode',       // 图片链接列名
  storeCodeColumnName: 'name',   // 门店ID列名（用于文件名）
  download: {
    concurrency: 50,                   // 并发数调低（减少服务器限流）
    timeout: 20000,                   // 超时延长至20秒
    retry: {
      maxTimes: 3,                    // 最大重试次数（失败后自动重试3次）
      initialDelay: 2000,             // 初始重试间隔（2秒）
      delayMultiplier: 1.5            // 重试间隔递增倍数（每次×1.5，避免频繁请求）
    },
    headers: {                        // 防盗链核心头（WeWork专用）
      'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36',
      'Referer': 'https://work.weixin.qq.com/',
      'Accept': 'image/png,image/jpeg,image/gif,image/svg+xml;image/webp,*/*;q=0.8'
    }
  },
  validation: {
    minFileSize: 100,                 // 最小图片大小（字节）：排除空文件/HTML拦截页
    imageHeaders: {                   // 图片文件头校验（确保是真实图片）
      png: [0x89, 0x50, 0x4E, 0x47],
      jpg: [0xFF, 0xD8, 0xFF],
      jpeg: [0xFF, 0xD8, 0xFF],
      gif: [0x47, 0x49, 0x46]
    }
  }
};

// 确保目录存在
async function ensureDirs() {
  await fsExtra.ensureDir(path.dirname(config.inputFile));
  await fsExtra.emptyDir(config.tempDir);
  await fsExtra.ensureDir(config.tempDir);
  await fsExtra.ensureDir(config.outputDir);
}

// 记录错误日志（含详细上下文）
function logError(message, task = null) {
  const timestamp = new Date().toISOString();
  const errorLogPath = path.join(config.outputDir, 'error_log.txt');
  // 追加详细信息（含门店ID和链接）
  let logMsg = `[${timestamp}] ${message}`;
  if (task) logMsg += ` | 门店ID: ${task.storeCode} | 链接: ${task.imageUrl}`;
  fs.appendFileSync(errorLogPath, `${logMsg}\n`);
  console.error(`❌ ${logMsg}`);
}

// 读取输入表格（保留原始数据，便于排查）
function readImageUrlsAndStoreCodesFromExcel() {
  try {
    if (!fs.existsSync(config.inputFile)) {
      throw new Error(`输入文件不存在：${config.inputFile}`);
    }

    const workbook = XLSX.readFile(config.inputFile);
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    // 过滤有效数据 + 保留原始行数据
    const validData = jsonData
      .map((item, index) => ({
        rowIndex: index + 2,            // 表格行号（从2开始，跳过表头）
        imageUrl: item[config.imageUrlColumnName]?.toString().trim(),
        storeCode: item[config.storeCodeColumnName]?.toString().trim(),
        rawItem: item                   // 原始数据，便于后续排查
      }))
      .filter(({ imageUrl, storeCode }) => {
        const isUrlValid = imageUrl && !['undefined', 'null'].includes(imageUrl.toLowerCase());
        const isStoreCodeValid = storeCode && !['undefined', 'null'].includes(storeCode.toLowerCase());
        return isUrlValid && isStoreCodeValid;
      });

    if (validData.length === 0) {
      throw new Error(`未读取到有效数据（需同时包含"${config.imageUrlColumnName}"和"${config.storeCodeColumnName}"列）`);
    }

    console.log(`📋 成功读取 ${validData.length} 条有效任务`);
    return validData;
  } catch (error) {
    logError(`读取表格失败：${error.message}`);
    throw error;
  }
}

// 1. 从响应头/文件头推断图片后缀
function getImageExtension(task, response = null, fileBuffer = null) {
  // 优先从响应头获取
  if (response) {
    const contentType = response.headers['content-type'] || '';
    const extMap = { 'image/png': '.png', 'image/jpeg': '.jpg', 'image/jpg': '.jpg', 'image/gif': '.gif' };
    if (extMap[contentType]) return extMap[contentType];
  }
  // 从文件头推断（下载后校验用）
  if (fileBuffer) {
    const header = fileBuffer.slice(0, 4);
    if (header[0] === 0x89 && header[1] === 0x50 && header[2] === 0x4E && header[3] === 0x47) return '.png';
    if (header[0] === 0xFF && header[1] === 0xD8 && header[2] === 0xFF) return '.jpg';
    if (header[0] === 0x47 && header[1] === 0x49 && header[2] === 0x46) return '.gif';
  }
  // 兜底：从链接获取
  return path.extname(task.imageUrl) || '.png';
}

// 2. 校验文件是否为有效图片（排除空文件/HTML拦截页）
function validateImageFile(task, filePath) {
  try {
    // 校验文件大小
    const stats = fs.statSync(filePath);
    if (stats.size < config.validation.minFileSize) {
      fs.unlinkSync(filePath); // 删除无效文件
      throw new Error(`文件过小（${stats.size}字节 < 最小${config.validation.minFileSize}字节），可能是拦截页`);
    }

    // 校验文件头（确保是真实图片）
    const fileBuffer = fs.readFileSync(filePath, { encoding: null, flag: 'r' });
    const ext = getImageExtension(task, null, fileBuffer);
    const expectedHeader = config.validation.imageHeaders[ext.slice(1)];
    if (!expectedHeader) throw new Error(`不支持的图片格式：${ext}`);

    // 比对文件头
    for (let i = 0; i < expectedHeader.length; i++) {
      if (fileBuffer[i] !== expectedHeader[i]) {
        fs.unlinkSync(filePath); // 删除无效文件
        throw new Error(`文件头不匹配（实际：${fileBuffer.slice(0, 4).join(',')} | 预期：${expectedHeader.join(',')}）`);
      }
    }

    // 若后缀不匹配，重命名文件（如链接无后缀但实际是jpg）
    const currentExt = path.extname(filePath);
    if (currentExt !== ext) {
      const newFilePath = filePath.replace(currentExt, ext);
      fs.renameSync(filePath, newFilePath);
      console.log(`ℹ️ 门店${task.storeCode}：文件后缀修正为${ext}（原${currentExt}）`);
    }

    return true;
  } catch (error) {
    throw new Error(`文件校验失败：${error.message}`);
  }
}

// 3. 单个图片下载（含自动重试+有效性校验）
async function downloadSingleImageWithRetry(task) {
  const { imageUrl, storeCode } = task;
  let retryCount = 0;

  // 循环重试：直到成功或达到最大重试次数
  while (retryCount < config.download.retry.maxTimes) {
    try {
      // 计算重试间隔（递增：2s → 3s → 4.5s...）
      if (retryCount > 0) {
        const delay = config.download.retry.initialDelay * Math.pow(config.download.retry.delayMultiplier, retryCount - 1);
        console.log(`ℹ️ 门店${storeCode}：第${retryCount}次重试，等待${Math.round(delay)}ms...`);
        await new Promise(resolve => setTimeout(resolve, delay));
      }

      // 发起请求（带防盗链头）
      const response = await axios({
        method: 'get',
        url: imageUrl,
        headers: config.download.headers,
        responseType: 'stream',
        timeout: config.download.timeout,
        withCredentials: true // 携带Cookie，应对部分登录态校验
      });

      // 校验响应状态和类型
      if (response.status !== 200) throw new Error(`服务器返回${response.status}状态码`);
      if (!response.headers['content-type']?.startsWith('image/')) {
        throw new Error(`响应非图片类型（Content-Type：${response.headers['content-type']}）`);
      }

      // 生成文件名并写入文件
      const ext = getImageExtension(task, response);
      const fileName = `${storeCode}${ext}`;
      const filePath = path.join(config.outputDir, fileName);
      await new Promise((resolve, reject) => {
        const writeStream = fs.createWriteStream(filePath);
        response.data.pipe(writeStream);

        writeStream.on('finish', () => resolve(filePath));
        writeStream.on('error', (err) => reject(new Error(`文件写入失败：${err.message}`)));
        response.data.on('error', (err) => reject(new Error(`响应流错误：${err.message}`)));
      });

      // 校验文件有效性（关键：排除无效文件）
      validateImageFile(task, filePath);

      console.log(`✅ 门店${storeCode}：下载成功（重试${retryCount}次）→ ${fileName}`);
      return { success: true, task, filePath };

    } catch (error) {
      retryCount++;
      const isLastRetry = retryCount >= config.download.retry.maxTimes;
      const errorMsg = isLastRetry 
        ? `最终失败（已重试${retryCount}次）：${error.message}` 
        : `重试${retryCount}次失败：${error.message}`;
      
      if (isLastRetry) logError(errorMsg, task); // 最后一次失败才记录日志
      else console.log(`⚠️ 门店${storeCode}：${errorMsg}`);
    }
  }

  // 达到最大重试次数仍失败
  return { success: false, task, error: `已达到最大重试次数（${config.download.retry.maxTimes}次）` };
}

// 4. 并发控制下载（稳定控制请求节奏）
async function downloadWithConcurrency(tasks) {
  const results = { success: [], fail: [] };
  const taskQueue = [...tasks];
  const running = [];

  while (taskQueue.length > 0 || running.length > 0) {
    // 控制并发数：未达上限且有任务时，启动新任务
    while (running.length < config.download.concurrency && taskQueue.length > 0) {
      const task = taskQueue.shift();
      const promise = downloadSingleImageWithRetry(task)
        .then(result => {
          // 分类记录结果
          if (result.success) results.success.push(result);
          else results.fail.push(result);
          running.splice(running.indexOf(promise), 1); // 从运行池移除
        });
      running.push(promise);
    }
    // 等待任一任务完成，避免循环空转
    if (running.length > 0) await Promise.race(running);
  }

  return results;
}

// 5. 失败任务二次处理（手动确认+重试）
async function handleFailedTasks(failedResults) {
  if (failedResults.length === 0) return { success: [], fail: [] };

  console.log(`\n🔴 存在 ${failedResults.length} 个下载失败的任务，开始二次处理...`);
  console.log(`📝 失败任务详情：`);
  failedResults.forEach((result, index) => {
    console.log(`  ${index + 1}. 门店ID：${result.task.storeCode} | 链接：${result.task.imageUrl} | 原因：${result.error}`);
  });

  // 询问用户是否二次重试（仅支持手动确认，避免自动重试过度）
  const readline = require('readline').createInterface({
    input: process.stdin,
    output: process.stdout
  });

  return new Promise(resolve => {
    readline.question(`\n是否对这些失败任务进行二次重试？（输入 y 重试，其他键跳过）：`, async (answer) => {
      readline.close();
      if (answer.toLowerCase() !== 'y') {
        console.log(`ℹ️ 已跳过二次重试，失败任务记录在 error_log.txt 中`);
        resolve({ success: [], fail: failedResults });
        return;
      }

      // 二次重试（使用相同逻辑，并发数更低）
      const tempConcurrency = Math.max(1, Math.floor(config.download.concurrency / 2));
      console.log(`\n🚀 开始二次重试（并发数：${tempConcurrency}）...`);
      const retryResults = await downloadWithConcurrency(failedResults.map(r => r.task));
      console.log(`\n二次重试完成：成功${retryResults.success.length}个，仍失败${retryResults.fail.length}个`);
      resolve(retryResults);
    });
  });
}

// 主函数：闭环保障流程
async function main() {
  try {
    await ensureDirs();
    console.log(`🚀 图片下载脚本启动（并发数：${config.download.concurrency} | 重试次数：${config.download.retry.maxTimes}次）`);

    // 1. 读取任务列表
    const allTasks = readImageUrlsAndStoreCodesFromExcel();

    // 2. 首次并发下载（含自动重试）
    console.log(`\n📥 开始首次下载（共 ${allTasks.length} 个任务）...`);
    const firstDownloadResults = await downloadWithConcurrency(allTasks);

    // 3. 二次处理失败任务（手动确认重试）
    const secondRetryResults = await handleFailedTasks(firstDownloadResults.fail);

    // 4. 最终统计（合并首次和二次结果）
    const finalSuccess = [...firstDownloadResults.success, ...secondRetryResults.success];
    const finalFail = secondRetryResults.fail;
    const successRate = (finalSuccess.length / allTasks.length * 100).toFixed(2);

    console.log(`\n🎉 所有下载流程完成！`);
    console.log(`📊 最终统计：`);
    console.log(`✅ 成功：${finalSuccess.length} 个（${successRate}%）`);
    console.log(`❌ 失败：${finalFail.length} 个`);
    console.log(`📁 图片保存路径：${config.outputDir}`);
    console.log(`📝 错误日志路径：${path.join(config.outputDir, 'error_log.txt')}`);

    // 输出最终失败任务（便于手动处理）
    if (finalFail.length > 0) {
      console.log(`\n🔴 最终失败任务清单（共${finalFail.length}个）：`);
      finalFail.forEach((result, index) => {
        console.log(`  ${index + 1}. 门店ID：${result.task.storeCode} | 表格行号：${result.task.rowIndex} | 链接：${result.task.imageUrl}`);
      });
    }

  } catch (error) {
    console.error(`\n💥 主流程中断：${error.message}`);
    logError(`主流程中断：${error.message}`);
  } finally {
    console.log(`\n📋 脚本执行结束`);
    process.exit(0);
  }
}

// 执行主函数
main().catch(error => {
  logError(`主函数未捕获错误：${error.message}`);
  process.exit(1);
});