/**
 * @对象    优惠券
 * @动作    导出
 * @风险    低
 * @来源    store-data-extractor/DaoChuQuanBuYouHuiQuan.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

// 导出互联商务中心表格

const axios = require("axios");
const ExcelJS = require("exceljs");
const fs = require("fs-extra");
const path = require("path");
const moment = require("moment");

// 配置
const config = {
  apiUrl: "https://vinci-api.feibing.tech/sc/v1/sellers/wwf8b0baf5a858670f/coupons",
  pageSize: 100,
  headers: {
    'accept': '*/*',
    'accept-language': 'zh-CN,zh;q=0.9',
    'authorization': 'bearer __VINCI_TOKEN__',
    'cookie': 'x-token=__VINCI_TOKEN__; acw_tc=0a0f705417675165710888741e36d3d0e9c21b28adc938fcbf3b0bc90a00d0',
    'origin': 'https://connect.feibing.tech',
    'referer': 'https://connect.feibing.tech/',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/143.0.0.0 Safari/537.36',
    'sec-ch-ua': '"Google Chrome";v="143", "Chromium";v="143", "Not A(Brand";v="24"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"macOS"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-site',
    'priority': 'u=1, i'
  },
  outputDir: path.join(__dirname, "../output/优惠券数据"),
  tempDir: path.join(__dirname, "../output/优惠券数据/temp"),
};

// 确保目录存在
async function initDirs() {
  await fs.ensureDir(config.outputDir);
  await fs.ensureDir(config.tempDir);
  await fs.emptyDir(config.tempDir);
}

// 格式化时间
function formatTime(timestamp) {
  if (!timestamp) return "";
  return moment(timestamp).format("YYYY-MM-DD HH:mm:ss");
}

// 获取单页数据
async function fetchPage(pageNum) {
  try {
    const response = await axios.get(config.apiUrl, {
      params: {
        pageNum: pageNum,
        pageSize: config.pageSize,
        moldId: "-",
        sort: "createdTime,desc",
        consumedTimeGte: "",
        consumedTimeLt: ""
      },
      headers: config.headers,
      timeout: 15000
    });
    
    if (response.data.code === 0) {
      return response.data;
    } else {
      console.error(`接口报错 (第 ${pageNum} 页): ${response.data.message}`);
      return null;
    }
  } catch (error) {
    console.error(`请求失败 (第 ${pageNum} 页): ${error.message}`);
    return null;
  }
}

// 保存单页Excel
async function savePageToExcel(data, pageNum) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(`优惠券_第${pageNum}页`);

  // 表头配置
  worksheet.columns = [
    { header: "券码", key: "code", width: 40 },
    { header: "ID", key: "id", width: 30 },
    { header: "生成时间", key: "createdTime", width: 20 },
    { header: "核销时间", key: "consumedTime", width: 20 },
    { header: "权益", key: "permissionDetail", width: 30 },
    { header: "状态名称", key: "statusName", width: 15 },
    { header: "状态码", key: "status", width: 15 },
    { header: "已绑定门店", key: "shopName", width: 30 },
    { header: "对应商户", key: "targetSellerName", width: 30 },
  ];

  // 添加数据
  data.forEach(item => {
    worksheet.addRow({
      code: item.code || "",
      id: item.id || "",
      createdTime: formatTime(item.createdTime),
      consumedTime: formatTime(item.consumedTime),
      permissionDetail: item.metadata?.permissionDetail || "",
      statusName: item.statusName || "",
      status: item.status || "",
      shopName: item.metadata?.shopName || "",
      targetSellerName: item.metadata?.targetSellerName || "",
    });
  });

  const filePath = path.join(config.tempDir, `优惠券_第${pageNum}页.xlsx`);
  await workbook.xlsx.writeFile(filePath);
  console.log(`✅ 已保存: 第 ${pageNum} 页`);
  return filePath;
}

// 合并所有文件
async function mergeAllFiles(filePaths) {
  const mergedWorkbook = new ExcelJS.Workbook();
  const mergedWorksheet = mergedWorkbook.addWorksheet("优惠券汇总");

  mergedWorksheet.columns = [
    { header: "券码", key: "code", width: 40 },
    { header: "ID", key: "id", width: 30 },
    { header: "生成时间", key: "createdTime", width: 20 },
    { header: "核销时间", key: "consumedTime", width: 20 },
    { header: "权益", key: "permissionDetail", width: 30 },
    { header: "状态名称", key: "statusName", width: 15 },
    { header: "状态码", key: "status", width: 15 },
    { header: "已绑定门店", key: "shopName", width: 30 },
    { header: "对应商户", key: "targetSellerName", width: 30 },
  ];

  console.log("\n🔗 开始汇总所有数据...");
  for (const filePath of filePaths) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.worksheets[0];

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) { // 跳过表头
        mergedWorksheet.addRow({
          code: row.getCell(1).value,
          id: row.getCell(2).value,
          createdTime: row.getCell(3).value,
          consumedTime: row.getCell(4).value,
          permissionDetail: row.getCell(5).value,
          statusName: row.getCell(6).value,
          status: row.getCell(7).value,
          shopName: row.getCell(8).value,
          targetSellerName: row.getCell(9).value,
        });
      }
    });
  }

  const finalPath = path.join(config.outputDir, `优惠券汇总_${Date.now()}.xlsx`);
  await mergedWorkbook.xlsx.writeFile(finalPath);
  console.log(`\n🎉 汇总完成！文件已保存至: ${finalPath}`);
  
  // 清理临时目录
  await fs.remove(config.tempDir);
}

// 主程序
async function main() {
  try {
    await initDirs();
    console.log("🚀 开始导出优惠券数据...");

    // 获取第一页确定总数
    const firstPage = await fetchPage(1);
    if (!firstPage) return;

    const total = firstPage.total;
    const totalPages = Math.ceil(total / config.pageSize);
    console.log(`📊 总计 ${total} 条数据，共 ${totalPages} 页`);

    const tempFiles = [];
    
    // 处理第一页
    const firstPath = await savePageToExcel(firstPage.data, 1);
    tempFiles.push(firstPath);

    // 循环获取剩余页面
    for (let page = 2; page <= totalPages; page++) {
      const pageData = await fetchPage(page);
      if (pageData && pageData.data) {
        const filePath = await savePageToExcel(pageData.data, page);
        tempFiles.push(filePath);
      }
      // 适当延迟防止频率限制
      await new Promise(resolve => setTimeout(resolve, 500));
    }

    await mergeAllFiles(tempFiles);

  } catch (error) {
    console.error("💥 程序崩溃:", error.message);
  }
}

main();

