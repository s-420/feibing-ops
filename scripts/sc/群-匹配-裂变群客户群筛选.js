/**
 * @对象    群
 * @动作    匹配
 * @风险    低
 * @来源    store-data-extractor/PiPeiLieBianQunHeKeHuQunShaiXuan.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const fsExtra = require('fs-extra');

// 匹配裂变群与客户群数据，并将门店名称单独列出来及判断包含关系的配置

const config = {
  inputDir: path.join(__dirname, '../input'),
  outputDir: path.join(__dirname, '../output/裂变群相关'),
  outputFile: `飞冰一点点-裂变群与客户群匹配结果_${new Date().toISOString().replace(/[:.]/g, '-').slice(0,19)}.xlsx`, // 输出文件名
  mainDataFile: '飞冰一点点-全部裂变群列表.xlsx',
  chatListFile: '飞冰一点点-全部客户群列表.xlsx',
  mainSheetName: '数据列表',
  chatSheetName: '活动列表',
  idColumn: 'id',
  nameColumn: 'name',
  shopColumn: '关联门店',
  chatIdListColumn: 'chat_id_list',
  roomBaseNameColumn: 'room_base_name', // 用于判断的列名
  shopExtractColumn: '提取的门店名称',  // 新列：提取的门店名称
  includeFlagColumn: '是否包含门店',    // 新列：是否包含标记
  deleteColumn: 'qr_code',
  idSeparators: [',', '，', ';', '；', ' ']
};

// 确保输出目录存在
async function ensureOutputDir() {
  await fsExtra.ensureDir(config.outputDir);
}

// 读取Excel文件数据（使用exceljs）
async function readExcelFile(filePath, sheetName) {
  try {
    console.log(`尝试读取文件: ${filePath}`);
    if (!fs.existsSync(filePath)) {
      throw new Error(`文件不存在，请检查路径是否正确`);
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) {
      throw new Error(`工作表"${sheetName}"不存在，请检查Excel中的工作表名`);
    }

    // 转换为JSON（包含表头）
    const data = [];
    const headers = {};
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) {
        // 读取表头
        row.eachCell((cell, colNumber) => {
          headers[colNumber] = cell.value;
        });
      } else {
        // 读取数据行
        const rowData = {};
        row.eachCell((cell, colNumber) => {
          rowData[headers[colNumber]] = cell.value;
        });
        data.push(rowData);
      }
    });

    console.log(`成功读取文件，共${data.length}条数据`);
    return data;
  } catch (error) {
    console.error(`读取Excel失败:`, error.message);
    throw error;
  }
}

// 分割ID列表
function splitIds(idListStr) {
  if (!idListStr) return [];
  
  let ids = [idListStr];
  config.idSeparators.forEach(separator => {
    ids = ids.flatMap(id => id.split(separator).map(item => item.trim()));
  });
  return ids.filter(id => id);
}

// 创建ID到信息的映射
function createIdToInfoMap(chatListData) {
  const map = new Map();
  chatListData.forEach(item => {
    const id = item[config.idColumn];
    if (!id) {
      console.warn(`客户群数据中存在无效的ID: ${JSON.stringify(item)}`);
      return;
    }
    map.set(id, {
      name: item[config.nameColumn] || '无名称',
      shop: item[config.shopColumn] || '无关联门店'
    });
  });
  console.log(`创建ID到信息的映射完成，共${map.size}条有效映射`);
  return map;
}

// 提取门店名称（从《》中提取）
function extractShopNames(formattedStr) {
  if (!formattedStr) return [];
  
  const shopNames = [];
  const regex = /《([^》]+)》/g;
  let match;
  
  while ((match = regex.exec(formattedStr)) !== null) {
    shopNames.push(match[1]);
  }
  
  return shopNames;
}

// 判断roomBaseName是否包含任何一个门店名称（严格匹配：完整包含核心名称才认为匹配）
function checkIncludeShop(roomBaseName, shopNames) {
  if (!roomBaseName || !shopNames || shopNames.length === 0) {
    return null; // 无法判断
  }
  
  // 1. 移除roomBaseName中的"店门"二字并处理
  let processedBaseName = String(roomBaseName)
    .replace(/点门/g, '')
    .toLowerCase()
    .replace(/[^\u4e00-\u9fa5a-zA-Z0-9]/g, '');
  
  return shopNames.some(shop => {
    // 处理门店名称，移除"店"字结尾和特殊符号
    let processedShop = String(shop)
      .toLowerCase()
      .replace(/店$/g, '') // 移除结尾的"店"字
      .replace(/[^\u4e00-\u9fa5a-zA-Z0-9]/g, ''); // 移除特殊符号
      
    // 严格匹配：群名称必须完整包含处理后的门店名称
    return processedBaseName.includes(processedShop);
  });
}

// 处理主数据
function processMainData(mainData, idToInfoMap) {
  return mainData.map((item, index) => {
    if (index % 100 === 0) {
      console.log(`已处理${index + 1}/${mainData.length}条裂变群数据`);
    }
    
    const newItem = { ...item };
    const chatIdListStr = item[config.chatIdListColumn] || '';
    const ids = splitIds(chatIdListStr);
    
    // 处理chat_id_list并提取门店名称
    if (ids.length > 0) {
      const formattedItems = [];
      const shopNames = [];
      
      ids.forEach(id => {
        const info = idToInfoMap.get(id);
        if (info) {
          formattedItems.push(`${info.name}    《${info.shop}》`);
          shopNames.push(info.shop);
        } else {
          formattedItems.push(`未知ID(${id})`);
        }
      });
      
      const separator = getMainSeparator(chatIdListStr);
      newItem[config.chatIdListColumn] = formattedItems.join(separator);
      newItem[config.shopExtractColumn] = shopNames.join(separator);
    } else {
      newItem[config.shopExtractColumn] = '';
    }
    
    // 判断是否包含门店名称
    const roomBaseName = newItem[config.roomBaseNameColumn] || '';
    const shopNames = extractShopNames(newItem[config.chatIdListColumn]);
    const includeResult = checkIncludeShop(roomBaseName, shopNames);
    
    newItem[config.includeFlagColumn] = includeResult === true ? '包含' : 
                                       includeResult === false ? '不包含' : '无法判断';
    
    return newItem;
  });
}

// 获取主要分隔符
function getMainSeparator(str) {
  if (!str) return ',';
  const countMap = {};
  config.idSeparators.forEach(sep => {
    countMap[sep] = (str.match(new RegExp(sep, 'g')) || []).length;
  });
  return Object.entries(countMap).sort((a, b) => b[1] - a[1])[0][0];
}

// 写入Excel文件（含样式）
async function writeNewExcelFile(data) {
  try {
    const outputPath = path.join(config.outputDir, config.outputFile);
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(config.mainSheetName);

    // 添加表头
    const headers = Object.keys(data[0]);
    worksheet.addRow(headers);
    // 表头样式（加粗）
    worksheet.getRow(1).font = { bold: true };

    // 添加数据行并设置样式
    data.forEach((rowData, rowIndex) => {
      const row = worksheet.addRow(Object.values(rowData));
      
      // 处理chat_id_list列的富文本（标红门店名称）
      const chatIdListIndex = headers.indexOf(config.chatIdListColumn);
      if (chatIdListIndex !== -1) {
        const cell = row.getCell(chatIdListIndex + 1); // 列索引从1开始
        const cellValue = cell.value || '';
        
        // 清除单元格内容，准备添加富文本
        cell.value = { richText: [] };
        
        // 匹配并添加富文本（门店名称标红）
        let lastIndex = 0;
        const regex = /《([^》]+)》/g;
        let match;
        
        while ((match = regex.exec(cellValue)) !== null) {
          // 添加括号前的文本
          if (match.index > lastIndex) {
            cell.value.richText.push({
              text: cellValue.substring(lastIndex, match.index),
              font: { color: { argb: 'FF000000' } } // 黑色
            });
          }
          // 添加左括号
          cell.value.richText.push({
            text: '《',
            font: { color: { argb: 'FF000000' } }
          });
          // 添加门店名称（红色）
          cell.value.richText.push({
            text: match[1],
            font: { color: { argb: 'FFFF0000' } } // 红色
          });
          // 添加右括号
          cell.value.richText.push({
            text: '》',
            font: { color: { argb: 'FF000000' } }
          });
          lastIndex = regex.lastIndex;
        }
        
        // 添加剩余文本
        if (lastIndex < cellValue.length) {
          cell.value.richText.push({
            text: cellValue.substring(lastIndex),
            font: { color: { argb: 'FF000000' } }
          });
        }
      }
      
      // 处理"是否包含门店"列的样式（不包含的标红）
      const includeFlagIndex = headers.indexOf(config.includeFlagColumn);
      if (includeFlagIndex !== -1) {
        const cell = row.getCell(includeFlagIndex + 1);
        if (cell.value === '不包含') {
          cell.font = { color: { argb: 'FFFF0000' } }; // 红色
        } else if (cell.value === '包含') {
          cell.font = { color: { argb: 'FF008000' } }; // 绿色
        }
      }
    });

    // 调整列宽
    worksheet.columns.forEach(column => {
      column.width = 25;
    });

    // 保存文件
    await workbook.xlsx.writeFile(outputPath);
    console.log(`✅ 处理完成，结果已保存到: ${outputPath}`);
    return outputPath;
  } catch (error) {
    console.error(`写入Excel失败:`, error.message);
    throw error;
  }
}

// 主函数
async function main() {
  try {
    console.log('开始处理裂变群与客户群数据匹配...');
    await ensureOutputDir();
    
    const mainDataPath = path.join(config.inputDir, config.mainDataFile);
    const chatListPath = path.join(config.inputDir, config.chatListFile);
    
    console.log('===== 读取裂变群列表 =====');
    const mainData = await readExcelFile(mainDataPath, config.mainSheetName);
    
    console.log('\n===== 读取客户群列表 =====');
    const chatListData = await readExcelFile(chatListPath, config.chatSheetName);
    
    const idToInfoMap = createIdToInfoMap(chatListData);
    console.log('\n===== 开始处理数据 =====');
    const processedData = processMainData(mainData, idToInfoMap);
    await writeNewExcelFile(processedData);
    
    console.log('\n🎉 所有数据处理完成！已添加门店单独列和包含判断列');
  } catch (error) {
    console.error('\n💥 处理中断:', error.message);
  }
}

main();