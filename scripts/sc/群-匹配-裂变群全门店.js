/**
 * @对象    群
 * @动作    匹配
 * @风险    低
 * @来源    store-data-extractor/PiPeiLieBianQunHeQuanMenDian.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const ExcelJS = require('exceljs');
const fs = require('fs-extra');
const path = require('path');

// 匹配全门店是否配置入群链接

async function matchShopsWithGroups() {
    // 定义文件路径
    const mainFilePath = path.join(__dirname, '../input/全门店-飞冰.xlsx');
    const groupFilePath = path.join(__dirname, '../input/飞冰统计-裂变群.xlsx');
    const outputFilePath = path.join(__dirname, '../output/门店匹配结果.xlsx');

    try {
        // 检查文件是否存在
        if (!await fs.pathExists(mainFilePath)) {
            throw new Error(`主体文件不存在: ${mainFilePath}`);
        }
        if (!await fs.pathExists(groupFilePath)) {
            throw new Error(`裂变群文件不存在: ${groupFilePath}`);
        }

        // 创建工作簿读取器
        const mainWorkbook = new ExcelJS.Workbook();
        const groupWorkbook = new ExcelJS.Workbook();

        // 读取文件
        await mainWorkbook.xlsx.readFile(mainFilePath);
        await groupWorkbook.xlsx.readFile(groupFilePath);

        // 获取第一个工作表
        const mainWorksheet = mainWorkbook.worksheets[0];
        const groupWorksheet = groupWorkbook.worksheets[0];

        // 收集所有裂变群中的shop_id
        const groupShopIds = new Set();
        let groupShopIdColumn = -1;

        // 查找裂变群表格中shop_id所在列
        groupWorksheet.getRow(1).eachCell((cell, colNumber) => {
            if (cell.value === 'shop_id') {
                groupShopIdColumn = colNumber;
            }
        });

        if (groupShopIdColumn === -1) {
            throw new Error('裂变群表格中未找到shop_id列');
        }

        // 收集所有shop_id
        for (let i = 2; i <= groupWorksheet.rowCount; i++) {
            const row = groupWorksheet.getRow(i);
            const shopId = row.getCell(groupShopIdColumn).value;
            if (shopId) {
                groupShopIds.add(shopId.toString()); // 转为字符串避免类型问题
            }
        }

        // 处理主体表格
        let mainShopIdColumn = -1;

        // 查找主体表格中shopId所在列
        mainWorksheet.getRow(1).eachCell((cell, colNumber) => {
            if (cell.value === 'shopId') {
                mainShopIdColumn = colNumber;
            }
        });

        if (mainShopIdColumn === -1) {
            throw new Error('主体表格中未找到shopId列');
        }

        // 添加"是否配置链接"列
        const newColumn = mainWorksheet.columnCount + 1;
        mainWorksheet.getRow(1).getCell(newColumn).value = '是否配置链接';
        mainWorksheet.getRow(1).getCell(newColumn).font = { bold: true };

        // 遍历主体表格行进行匹配
        for (let i = 2; i <= mainWorksheet.rowCount; i++) {
            const row = mainWorksheet.getRow(i);
            const shopId = row.getCell(mainShopIdColumn).value;
            
            // 判断是否匹配
            const hasLink = shopId ? groupShopIds.has(shopId.toString()) : false;
            row.getCell(newColumn).value = hasLink ? '是' : '否';

            // 设置单元格样式
            row.getCell(newColumn).alignment = { horizontal: 'center' };
        }

        // 确保输出目录存在
        await fs.ensureDir(path.dirname(outputFilePath));

        // 保存结果
        await mainWorkbook.xlsx.writeFile(outputFilePath);
        console.log(`匹配完成，结果已保存至: ${outputFilePath}`);

    } catch (error) {
        console.error('处理过程中出错:', error.message);
        process.exit(1);
    }
}

// 执行匹配函数
matchShopsWithGroups();