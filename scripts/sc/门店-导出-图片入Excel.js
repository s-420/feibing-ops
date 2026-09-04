/**
 * @对象    门店
 * @动作    导出
 * @风险    低
 * @来源    store-data-extractor/DaoChuTuPianChaRuExcel.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

// 图片下载并嵌入Excel脚本（并发优化版）
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const https = require('https');
const url = require('url');

// ===================== 核心配置 =====================
const INPUT_DIR = path.join(__dirname, '../input');
const OUTPUT_DIR = path.join(__dirname, '../output');
const IMAGE_SAVE_DIR = path.join(OUTPUT_DIR, '生动化图片库');
const INPUT_FILE_NAME = '生动化.xlsx';
const INPUT_FILE_PATH = path.join(INPUT_DIR, INPUT_FILE_NAME);
const OUTPUT_FILE_NAME = `生动化_含图片_${new Date().toLocaleDateString().replace(/\//g, '-')}.xlsx`;
const OUTPUT_FILE_PATH = path.join(OUTPUT_DIR, OUTPUT_FILE_NAME);
const CONCURRENT_LIMIT = 10; // 并发数限制（建议5-20，根据服务器承受能力调整）

// ===================== 初始化文件夹 =====================
const initDir = () => {
  if (!fs.existsSync(INPUT_DIR)) {
    console.error(`❌ 输入文件夹 ${INPUT_DIR} 不存在，请创建后放入${INPUT_FILE_NAME}文件`);
    process.exit(1);
  }
  if (!fs.existsSync(INPUT_FILE_PATH)) {
    console.error(`❌ 未在input文件夹找到 ${INPUT_FILE_NAME} 文件，请检查文件名`);
    process.exit(1);
  }

  [OUTPUT_DIR, IMAGE_SAVE_DIR].forEach(dir => {
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
      console.log(`✅ 创建文件夹：${dir}`);
    }
  });
};

// ===================== 工具函数：清理URL =====================
const cleanImageUrl = (imgUrl) => {
  if (!imgUrl || typeof imgUrl !== 'string') return '';
  
  let cleanUrl = imgUrl.trim()
    .replace(/\u00A0/g, '') // 全角空格
    .replace(/\u200B/g, '') // 零宽空格
    .replace(/[\x00-\x1F\x7F]/g, ''); // 控制字符
  
  try {
    const parsedUrl = new URL(cleanUrl);
    parsedUrl.pathname = encodeURI(parsedUrl.pathname);
    cleanUrl = parsedUrl.toString();
  } catch (e) {
    console.warn(`⚠️ URL格式不标准，已尝试清理：${imgUrl}`);
  }
  
  return cleanUrl;
};

// ===================== 工具函数：处理文件名 =====================
const getSafeFilename = (filename) => {
  return filename.replace(/[\/:*?"<>|]/g, '-').replace(/\s+/g, '_');
};

// ===================== 工具函数：单个图片下载 =====================
const downloadSingleImage = async (task) => {
  const { imgUrl, rowIndex, colName, rowId } = task;
  const cleanUrl = cleanImageUrl(imgUrl);
  
  if (!cleanUrl) {
    return { ...task, success: false, message: '图片路径为空或无效', path: '', filename: '' };
  }

  // 提取文件名
  const urlParts = cleanUrl.split('/');
  let filename = urlParts[urlParts.length - 1] || `img_${rowId}_${colName}_${Date.now()}.jpg`;
  filename = getSafeFilename(filename);
  const savePath = path.join(IMAGE_SAVE_DIR, filename);

  // 模拟浏览器请求配置
  const parsedUrl = new URL(cleanUrl);
  const requestOptions = {
    hostname: parsedUrl.hostname,
    path: parsedUrl.pathname + parsedUrl.search,
    method: 'GET',
    headers: {
      'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36',
      'Referer': parsedUrl.origin + '/',
      'Accept': 'image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8',
      'Accept-Language': 'zh-CN,zh;q=0.9',
      'Connection': 'keep-alive'
    },
    rejectUnauthorized: false
  };

  return new Promise((resolve) => {
    const request = https.request(requestOptions, (response) => {
      // 跟随重定向
      if (response.statusCode >= 300 && response.statusCode < 400 && response.headers.location) {
        downloadSingleImage({ ...task, imgUrl: response.headers.location }).then(resolve);
        return;
      }

      if (response.statusCode !== 200) {
        resolve({ 
          ...task, 
          success: false, 
          message: `HTTP错误，状态码: ${response.statusCode}`, 
          path: '', 
          filename: '' 
        });
        return;
      }

      const fileStream = fs.createWriteStream(savePath);
      response.pipe(fileStream);

      fileStream.on('finish', () => {
        fileStream.close();
        resolve({ 
          ...task, 
          success: true, 
          message: '下载成功', 
          path: savePath, 
          filename 
        });
      });

      fileStream.on('error', (err) => {
        if (fs.existsSync(savePath)) fs.unlinkSync(savePath);
        resolve({ 
          ...task, 
          success: false, 
          message: `文件写入失败: ${err.message}`, 
          path: '', 
          filename: '' 
        });
      });
    });

    // 15秒超时
    request.setTimeout(15000, () => {
      request.abort();
      resolve({ 
        ...task, 
        success: false, 
        message: '下载超时（15秒）', 
        path: '', 
        filename: '' 
      });
    });

    request.on('error', (err) => {
      resolve({ 
        ...task, 
        success: false, 
        message: `请求失败: ${err.message}`, 
        path: '', 
        filename: '' 
      });
    });

    request.end();
  });
};

// ===================== 工具函数：分批并发下载 =====================
const batchConcurrentDownload = async (tasks) => {
  const totalTasks = tasks.length;
  let completedTasks = 0;
  let results = [];
  console.log(`📥 开始并发下载，共 ${totalTasks} 个图片任务，并发数：${CONCURRENT_LIMIT}`);

  // 分批执行
  const batches = [];
  for (let i = 0; i < totalTasks; i += CONCURRENT_LIMIT) {
    batches.push(tasks.slice(i, i + CONCURRENT_LIMIT));
  }

  // 处理每一批
  for (const [batchIndex, batch] of batches.entries()) {
    console.log(`\n🚀 处理第 ${batchIndex + 1}/${batches.length} 批任务（${batch.length} 个图片）`);
    
    // 批量执行当前批次
    const batchResults = await Promise.allSettled(
      batch.map(task => downloadSingleImage(task))
    );

    // 收集结果（过滤rejected状态）
    const validResults = batchResults
      .filter(result => result.status === 'fulfilled')
      .map(result => result.value);
    
    results = [...results, ...validResults];
    
    // 更新进度
    completedTasks += validResults.length;
    console.log(`✅ 第 ${batchIndex + 1} 批完成，累计完成 ${completedTasks}/${totalTasks} 个任务`);
  }

  return results;
};

// ===================== 工具函数：嵌入图片到Excel =====================
const insertImageToExcel = async (worksheet, rowIndex, colIndex, imgPath) => {
  if (!imgPath || !fs.existsSync(imgPath)) {
    return false;
  }

  try {
    const imageBuffer = fs.readFileSync(imgPath);
    const imgExt = path.extname(imgPath).toLowerCase();
    let imageType = 'jpeg';
    if (imgExt === '.png') imageType = 'png';
    if (imgExt === '.gif') imageType = 'gif';

    const imageId = worksheet.workbook.addImage({
      buffer: imageBuffer,
      extension: imageType
    });

    worksheet.addImage(imageId, {
      tl: { col: colIndex - 1, row: rowIndex - 1 },
      br: { col: colIndex, row: rowIndex + 3 },
      editAs: 'oneCell'
    });

    worksheet.getRow(rowIndex).height = 20;
    worksheet.getColumn(colIndex).width = 30;
    return true;
  } catch (err) {
    console.error(`⚠️  嵌入图片失败 ${imgPath}:`, err.message);
    return false;
  }
};

// ===================== 主函数 =====================
const main = async () => {
  try {
    // 1. 初始化文件夹
    initDir();
    console.log('📋 开始读取Excel文件...');

    // 2. 读取Excel并收集下载任务
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(INPUT_FILE_PATH);
    const worksheet = workbook.worksheets[0];
    console.log(`✅ 成功读取工作表：${worksheet.name}`);

    // 查找Img/TaskImg列
    let imgColIndex = -1;
    let taskImgColIndex = -1;
    worksheet.getRow(1).eachCell((cell, colIndex) => {
      const cellValue = (cell.value || '').toString().trim().toLowerCase();
      if (cellValue === 'img') imgColIndex = colIndex;
      if (cellValue === 'taskimg') taskImgColIndex = colIndex;
    });

    if (imgColIndex === -1 && taskImgColIndex === -1) {
      console.error('❌ 未在表头找到 Img 或 TaskImg 列，请检查表格列名');
      process.exit(1);
    }
    console.log(`📌 找到目标列 - Img列：${imgColIndex || '未找到'}，TaskImg列：${taskImgColIndex || '未找到'}`);

    // 收集所有下载任务
    const downloadTasks = [];
    const totalRows = worksheet.rowCount;
    const dataRows = totalRows - 1;

    for (let rowIndex = 2; rowIndex <= totalRows; rowIndex++) {
      const row = worksheet.getRow(rowIndex);
      const rowId = rowIndex - 1; // 数据行ID（从1开始）

      // 添加Img列任务
      if (imgColIndex !== -1) {
        const imgUrl = (row.getCell(imgColIndex).value || '').toString().trim();
        if (imgUrl) {
          downloadTasks.push({
            imgUrl,
            rowIndex,
            colIndex: imgColIndex,
            colName: 'Img',
            rowId
          });
        }
      }

      // 添加TaskImg列任务
      if (taskImgColIndex !== -1) {
        const taskImgUrl = (row.getCell(taskImgColIndex).value || '').toString().trim();
        if (taskImgUrl) {
          downloadTasks.push({
            imgUrl: taskImgUrl,
            rowIndex,
            colIndex: taskImgColIndex,
            colName: 'TaskImg',
            rowId
          });
        }
      }
    }

    if (downloadTasks.length === 0) {
      console.log('📭 未找到需要下载的图片URL，直接保存原表格');
      await workbook.xlsx.writeFile(OUTPUT_FILE_PATH);
      console.log(`🎉 新表格已保存到：${OUTPUT_FILE_PATH}`);
      process.exit(0);
    }

    // 3. 并发下载所有图片
    const downloadResults = await batchConcurrentDownload(downloadTasks);

    // 4. 处理下载结果（嵌入图片或标记失败）
    console.log('\n📥 开始处理图片嵌入Excel...');
    let imgSuccessCount = 0;
    let taskImgSuccessCount = 0;

    for (const result of downloadResults) {
      const { rowIndex, colIndex, colName, success, path: imgPath, message } = result;
      const cell = worksheet.getRow(rowIndex).getCell(colIndex);

      if (success) {
        // 嵌入图片
        await insertImageToExcel(worksheet, rowIndex, colIndex, imgPath);
        if (colName === 'Img') imgSuccessCount++;
        if (colName === 'TaskImg') taskImgSuccessCount++;
      } else {
        // 标记失败原因
        cell.value = `下载失败：${message}`;
      }
    }

    // 5. 优化Excel格式
    worksheet.columns.forEach((column, index) => {
      const colIndex = index + 1;
      if (colIndex !== imgColIndex && colIndex !== taskImgColIndex) {
        let maxWidth = (column.header || '').length + 2;
        worksheet.getColumn(colIndex).eachCell((cell) => {
          if (cell.value && typeof cell.value === 'string') {
            maxWidth = Math.min(Math.max(maxWidth, cell.value.length + 2), 40);
          }
        });
        worksheet.getColumn(colIndex).width = Math.max(maxWidth, 12);
      }
    });

    // 6. 保存最终表格
    await workbook.xlsx.writeFile(OUTPUT_FILE_PATH);
    console.log(`\n🎉 新表格已保存到：${OUTPUT_FILE_PATH}`);

    // 7. 输出统计信息
    console.log('\n📊 最终统计：');
    console.log(`- 总下载任务数：${downloadTasks.length}`);
    console.log(`- 成功下载数：${downloadResults.filter(r => r.success).length}`);
    console.log(`- Img列：成功 ${imgSuccessCount} 张，失败 ${(imgColIndex !== -1 ? dataRows : 0) - imgSuccessCount} 张`);
    console.log(`- TaskImg列：成功 ${taskImgSuccessCount} 张，失败 ${(taskImgColIndex !== -1 ? dataRows : 0) - taskImgSuccessCount} 张`);
    console.log(`- 图片保存路径：${IMAGE_SAVE_DIR}`);
    console.log(`- 并发数：${CONCURRENT_LIMIT}（可在脚本顶部配置调整）`);

  } catch (error) {
    console.error('\n❌ 脚本执行失败：', error.message);
    process.exit(1);
  }
};

// 启动脚本
main();