/**
 * @对象    门店
 * @动作    匹配
 * @风险    低
 * @来源    store-data-extractor/BiDuiTuPianMingChengYuShuJu.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const fs = require('fs').promises;
const path = require('path');
const xlsx = require('xlsx');
const { existsSync, mkdirSync } = require('fs');

// 获取对应的图片ID（去扩展名）与Excel中的三方ID进行比对，删除不匹配的图片
// 适用于：图片命名为三方ID（去扩展名），Excel中有三方ID列
// 注意：删除操作不可逆，首次测试请务必将 ENABLE_DELETE 设为 false，仅预览不执行删除

// -------------------------- 路径配置（根据实际项目调整）--------------------------
const ROOT_PATH = path.join(__dirname, '../input'); // input文件夹路径（脚本上级目录下的input）
const IMAGE_FOLDER = path.join(ROOT_PATH, 'POS09'); // 图片存放目录（input/POS09）
const EXCEL_FILE = path.join(ROOT_PATH, '全门店已上架数据.xlsx'); // Excel文件路径（input下的Excel）
const LOG_FILE = path.join(__dirname, '../output/delete_log.txt'); // 文本日志路径（脚本同级目录）
const OUTPUT_FOLDER = path.join(__dirname, '../output'); // Excel导出目录（与input同级的output）
const EXPORT_EXCEL_NAME = `删除图片ID列表_${new Date().getTime()}.xlsx`; // 导出Excel名（带时间戳防重复）
const EXCEL_SHEET_NAME = 'Sheet1'; // Excel工作表名
const EXCEL_COLUMN_NAME = '三方ID'; // Excel中用于比对的列名
const ENABLE_DELETE = true; // 安全开关：true=执行删除，false=仅预览（首次测试建议设false）
// -------------------------------------------------------------------------------

// 支持的图片格式（可根据需求补充）
const IMAGE_EXTENSIONS = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'];

/**
 * 检查文件是否为目标图片格式
 * @param {string} filename - 文件名
 * @returns {boolean} 是否为图片
 */
function isImageFile(filename) {
  const ext = path.extname(filename).toLowerCase();
  return IMAGE_EXTENSIONS.includes(ext);
}

/**
 * 读取POS09文件夹中的所有图片文件（含完整路径和ID）
 * @returns {Array} 图片信息数组（含fullName/nameWithoutExt/path/ext）
 */
async function readImageFiles() {
  try {
    // 校验POS09文件夹是否存在
    if (!existsSync(IMAGE_FOLDER)) {
      console.error(`❌ 错误：POS09文件夹不存在，路径：${IMAGE_FOLDER}`);
      return [];
    }

    const files = await fs.readdir(IMAGE_FOLDER);
    const imageFiles = [];

    for (const file of files) {
      const filePath = path.join(IMAGE_FOLDER, file);
      const stats = await fs.stat(filePath);

      // 仅保留文件且为目标图片格式
      if (stats.isFile() && isImageFile(file)) {
        const fileNameWithoutExt = path.parse(file).name; // 图片ID（去扩展名）
        imageFiles.push({
          fullName: file, // 完整文件名（含扩展名）
          nameWithoutExt: fileNameWithoutExt, // 用于比对的图片ID
          path: filePath, // 文件绝对路径
          ext: path.extname(file).toLowerCase() // 文件扩展名
        });
      }
    }

    console.log(`✅ 从POS09文件夹读取到 ${imageFiles.length} 张有效图片`);
    return imageFiles;
  } catch (err) {
    console.error(`❌ 读取图片文件夹失败：${err.message}`);
    return [];
  }
}

/**
 * 读取Excel中的三方ID（去重+数据清洗）
 * @returns {Array} 去重后的三方ID数组
 */
function readExcelThirdPartyIds() {
  try {
    // 校验Excel文件是否存在
    if (!existsSync(EXCEL_FILE)) {
      console.error(`❌ 错误：Excel文件不存在，路径：${EXCEL_FILE}`);
      return [];
    }

    // 读取Excel并转换为JSON
    const workbook = xlsx.readFile(EXCEL_FILE);
    const worksheet = workbook.Sheets[EXCEL_SHEET_NAME];

    // 校验工作表是否存在
    if (!worksheet) {
      console.error(`❌ 错误：Excel中不存在工作表 ${EXCEL_SHEET_NAME}`);
      return [];
    }

    const jsonData = xlsx.utils.sheet_to_json(worksheet);

    // 提取三方ID并清洗（去重、去空格、处理数字格式）
    const thirdPartyIds = [...new Set(
      jsonData.map(item => {
        const rawId = item[EXCEL_COLUMN_NAME];
        if (rawId === undefined || rawId === null) return null;

        // 清洗逻辑：转字符串、去空格、处理科学计数法
        let cleanId = String(rawId).trim();
        if (!isNaN(Number(cleanId)) && cleanId.includes('E')) {
          cleanId = Number(cleanId).toString(); // 解决科学计数法（如 5.926E4 → 59260）
        }

        return cleanId || null;
      }).filter(Boolean) // 过滤空值
    )];

    console.log(`✅ 从Excel读取到 ${thirdPartyIds.length} 个去重后的三方ID`);
    return thirdPartyIds;
  } catch (err) {
    console.error(`❌ 读取Excel失败：${err.message}`);
    return [];
  }
}

/**
 * 记录删除日志到文本文件（追加模式）
 * @param {Array} deletedFiles - 已删除的图片信息数组
 */
async function logDeletedFiles(deletedFiles) {
  if (deletedFiles.length === 0) return;

  const timestamp = new Date().toISOString();
  let logContent = `[${timestamp}] 本次删除 ${deletedFiles.length} 个文件：\n`;
  deletedFiles.forEach(file => {
    logContent += `- 文件名：${file.fullName} | 图片ID：${file.nameWithoutExt}\n`;
  });
  logContent += '----------------------------------------\n\n';

  try {
    await fs.appendFile(LOG_FILE, logContent);
    console.log(`✅ 文本日志已保存到：${LOG_FILE}`);
  } catch (err) {
    console.error(`❌ 写入文本日志失败：${err.message}`);
  }
}

/**
 * 导出删除的图片列表到Excel（output文件夹下）
 * @param {Array} deletedFiles - 已删除的图片信息数组
 */
async function exportDeletedToExcel(deletedFiles) {
  if (deletedFiles.length === 0) return;

  try {
    // 自动创建output文件夹（多级目录也支持）
    if (!existsSync(OUTPUT_FOLDER)) {
      mkdirSync(OUTPUT_FOLDER, { recursive: true });
      console.log(`✅ 自动创建output文件夹：${OUTPUT_FOLDER}`);
    }

    // 整理Excel数据（表头+内容）
    const excelData = [
      ['序号', '图片ID（去扩展名）', '完整文件名', '文件格式', '删除时间'],
      ...deletedFiles.map((file, index) => [
        index + 1,
        file.nameWithoutExt,
        file.fullName,
        file.ext,
        new Date().toLocaleString() // 本地时间（如 2024/5/20 15:30:00）
      ])
    ];

    // 创建Excel工作簿和工作表
    const workbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.aoa_to_sheet(excelData);

    // 优化列宽（避免内容截断）
    worksheet['!cols'] = [
      { wch: 6 },  // 序号列宽
      { wch: 20 }, // 图片ID列宽
      { wch: 35 }, // 完整文件名列宽
      { wch: 12 }, // 文件格式列宽
      { wch: 22 }  // 删除时间列宽
    ];

    // 保存Excel文件
    xlsx.utils.book_append_sheet(workbook, worksheet, '删除图片列表');
    const exportPath = path.join(OUTPUT_FOLDER, EXPORT_EXCEL_NAME);
    xlsx.writeFile(workbook, exportPath);

    console.log(`✅ Excel删除列表已导出到：${exportPath}`);
  } catch (err) {
    console.error(`❌ 导出Excel失败：${err.message}`);
  }
}

/**
 * 执行不匹配图片的删除（含二次确认）
 * @param {Array} unmatchedImages - 不匹配的图片信息数组
 * @returns {Array} 实际删除成功的图片信息数组
 */
async function deleteUnmatchedImages(unmatchedImages) {
  if (unmatchedImages.length === 0) {
    console.log('✅ 没有需要删除的不匹配图片');
    return [];
  }

  // 预览待删除列表（仅显示前20个，避免输出过长）
  console.log(`\n⚠️  待删除不匹配图片共 ${unmatchedImages.length} 个（前20个预览）：`);
  unmatchedImages.slice(0, 20).forEach((img, index) => {
    console.log(`${index + 1}. 文件名：${img.fullName} | 图片ID：${img.nameWithoutExt}`);
  });
  if (unmatchedImages.length > 20) {
    console.log(`... 剩余 ${unmatchedImages.length - 20} 个待删除图片未显示`);
  }

  // 命令行二次确认（防止误操作）
  const readline = require('readline').createInterface({
    input: process.stdin,
    output: process.stdout
  });

  return new Promise((resolve) => {
    readline.question(`\n⚠️  确认要删除这 ${unmatchedImages.length} 个图片吗？（输入 y 确认，其他键取消）：`, async (answer) => {
      readline.close();

      // 取消删除
      if (answer.toLowerCase() !== 'y') {
        console.log('✅ 已取消删除操作');
        resolve([]);
        return;
      }

      const deletedFiles = [];
      let successCount = 0;
      let failCount = 0;

      // 执行删除（仅当ENABLE_DELETE为true时）
      if (ENABLE_DELETE) {
        console.log('\n🔄 开始执行删除...');
        for (const img of unmatchedImages) {
          try {
            await fs.unlink(img.path);
            console.log(`✅ 已删除：${img.fullName}（ID：${img.nameWithoutExt}）`);
            successCount++;
            deletedFiles.push(img);
          } catch (err) {
            console.error(`❌ 删除失败 ${img.fullName}（ID：${img.nameWithoutExt}）：${err.message}`);
            failCount++;
          }
        }

        console.log(`\n📊 删除完成：成功 ${successCount} 个 | 失败 ${failCount} 个`);
        resolve(deletedFiles);
      } else {
        // 安全开关关闭，仅预览不删除
        console.log('\nℹ️  删除功能未启用（ENABLE_DELETE=false），未执行实际删除');
        resolve([]);
      }
    });
  });
}

/**
 * 主函数：串联所有逻辑
 */
async function main() {
  console.log('========================================');
  console.log('🔍 开始执行图片比对删除脚本');
  console.log('========================================');

  // 1. 读取图片和Excel数据
  const imageFiles = await readImageFiles();
  if (imageFiles.length === 0) return;

  const thirdPartyIds = readExcelThirdPartyIds();
  if (thirdPartyIds.length === 0) return;

  // 2. 筛选不匹配的图片（图片ID不在三方ID列表中）
  const unmatchedImages = imageFiles.filter(img => 
    !thirdPartyIds.includes(img.nameWithoutExt)
  );
  console.log(`\n🔍 筛选出不匹配的图片数量：${unmatchedImages.length} 个`);

  // 3. 执行删除、日志记录、Excel导出
  const deletedFiles = await deleteUnmatchedImages(unmatchedImages);
  if (deletedFiles.length > 0) {
    await logDeletedFiles(deletedFiles);
    await exportDeletedToExcel(deletedFiles);
  }

  console.log('\n========================================');
  console.log('📝 脚本执行完成');
  console.log('========================================');
}

// 启动脚本
main();