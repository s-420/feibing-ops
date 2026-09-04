/**
 * @对象    门店
 * @动作    导出
 * @风险    低
 * @来源    store-data-extractor/DaoChuShannChuJiLu.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

require("dotenv").config();
const axios = require("axios");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const moment = require("moment");

// 配置项
const config = {
  baseUrl: "https://scrm.feibing.tech",
  endpoint: "/wework-scrm/del_remind/del_list",
  headers: {
    accept: "application/json",
    "accept-language": "zh-CN,zh;q=0.9",
    "content-type": "application/json;charset=UTF-8",
    origin: "https://scrm.feibing.tech",
    priority: "u=1, i",
    referer: "https://scrm.feibing.tech/",
    "sec-ch-ua":
      '"Chromium";v="142", "Google Chrome";v="142", "Not_A Brand";v="99"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"macOS"',
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-origin",
    "user-agent":
      "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36",
    cookie:
      "_bl_uid=ab3b0489-1a0b-40e0-b8a3-5f5685e318fb; _clck=c2po6l%5E2%5Eg47%5E0%5E2259; _clsk=17ulvqr%5E1773038561268%5E10%5E1%5Ef.clarity.ms%2Fcollect; x-token=__VINCI_TOKEN__",
  },
  requestBody: {
    pageSize: 100,
    addTime: {
      startTime: 1785513600000, 
      endTime: 1788191999999, 
    },
    filterInherit: true
  },
  outputDir: path.join(__dirname, "../output/删人记录"),
  concurrency: 10, // 同时执行的页面任务数
  startPage: 1, // 从第几页开始导出
};

// 确保输出目录存在
const ensureDirExists = (dirPath) => {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
    console.log(`创建输出目录: ${dirPath}`);
  }
};

// 格式化时间戳
const formatTime = (timestamp) => {
  if (!timestamp) return "";
  return moment(timestamp).format("YYYY-MM-DD HH:mm:ss");
};

// 创建Excel文件
const createExcelFile = async (data, filePath) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("删人记录");

  // 设置表头
  worksheet.columns = [
    { header: "删除客户", key: "name", width: 20 },
    { header: "操作人", key: "staffUserId", width: 30 },
    { header: "删除时间", key: "lostTime", width: 20 },
    { header: "添加好友时间", key: "createTime", width: 20 },
  ];

  // 添加数据
  data.forEach((item) => {
    worksheet.addRow({
      name: item.name || "",
      staffUserId: item.staffUserId || "",
      lostTime: formatTime(item.lostTime),
      createTime: formatTime(item.createTime),
    });
  });

  // 保存文件
  await workbook.xlsx.writeFile(filePath);
  console.log(`Excel文件已保存: ${filePath}`);
};

// 发送请求获取数据
const fetchData = async (page) => {
  try {
    const response = await axios.post(
      `${config.baseUrl}${config.endpoint}`,
      { ...config.requestBody, page },
      { headers: config.headers }
    );

    if (response.data.success && response.data.code === 200) {
      return response.data.data;
    } else {
      console.error(`请求失败 (第 ${page} 页):`, response.data.message);
      return null;
    }
  } catch (error) {
    console.error(`请求异常 (第 ${page} 页):`, error.message);
    return null;
  }
};

// 使用固定数量的 worker 处理任务：任意任务完成后，立即领取下一页
const runPagePool = async (pages, concurrency, worker) => {
  let nextIndex = 0;

  const runWorker = async () => {
    while (nextIndex < pages.length) {
      const currentIndex = nextIndex;
      nextIndex += 1;
      await worker(pages[currentIndex]);
    }
  };

  const workerCount = Math.min(concurrency, pages.length);
  await Promise.all(
    Array.from({ length: workerCount }, () => runWorker())
  );
};

// 主函数
const main = async () => {
  try {
    // 确保输出目录存在
    ensureDirExists(config.outputDir);

    // 先获取第一页数据，确定总页数
    console.log("正在获取第一页数据...");
    const firstPageData = await fetchData(1);
    if (!firstPageData) {
      console.error("获取第一页数据失败，程序终止");
      return;
    }

    const { totalPage } = firstPageData;
    console.log(`总页数: ${totalPage}, 总记录数: ${firstPageData.total}`);

    const startPage = config.startPage || 1;

    // 处理第一页数据
    if (startPage === 1) {
      const firstPageFilePath = path.join(
        config.outputDir,
        `删人记录_第1页.xlsx`
      );
      await createExcelFile(firstPageData.list, firstPageFilePath);
    } else {
      console.log(`跳过第 1 页，从第 ${startPage} 页开始处理...`);
    }

    // 使用固定并发池处理剩余页面，不再等待整批任务全部完成
    const loopStartPage = Math.max(2, startPage);
    const pages = Array.from(
      { length: Math.max(0, totalPage - loopStartPage + 1) },
      (_, index) => loopStartPage + index
    );

    if (pages.length > 0) {
      console.log(
        `\n正在处理第 ${loopStartPage} 到 ${totalPage} 页（固定并发 ${config.concurrency}）...`
      );

      await runPagePool(pages, config.concurrency, async (page) => {
        console.log(`开始处理第 ${page} 页...`);
        const pageData = await fetchData(page);

        if (!pageData) {
          console.warn(`第 ${page} 页数据获取失败`);
          return;
        }

        const filePath = path.join(
          config.outputDir,
          `删人记录_第${page}页.xlsx`
        );
        await createExcelFile(pageData.list, filePath);
        console.log(`第 ${page} 页处理完成，继续补充下一页`);
      });
    } else {
      console.log("没有需要继续处理的分页");
    }

    // 导出汇总Excel
    const summaryFilePath = path.join(config.outputDir, `删人记录_汇总.xlsx`);
    console.log(`\n正在读取目录下所有历史文件并生成汇总文件...`);
    
    // 扫描目录下所有的分页文件
    const files = fs.readdirSync(config.outputDir)
      .filter(f => f.startsWith('删人记录_第') && f.endsWith('.xlsx'))
      .sort((a, b) => {
        const numA = parseInt(a.match(/\d+/)[0]);
        const numB = parseInt(b.match(/\d+/)[0]);
        return numA - numB;
      });

    let finalSummaryData = [];
    for (const file of files) {
      const filePath = path.join(config.outputDir, file);
      try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const worksheet = workbook.worksheets[0];
        
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber > 1) { // 跳过表头
            finalSummaryData.push({
              name: row.getCell(1).value,
              staffUserId: row.getCell(2).value,
              lostTime: row.getCell(3).value,
              createTime: row.getCell(4).value
            });
          }
        });
        console.log(`已读取: ${file}`);
      } catch (e) {
        console.error(`读取文件 ${file} 失败:`, e.message);
      }
    }

    console.log(`\n合并完成，总记录数: ${finalSummaryData.length}`);
    await createExcelFile(finalSummaryData, summaryFilePath);

    console.log("\n所有任务完成！");
    console.log(`- 单页文件目录: ${config.outputDir}`);
    console.log(`- 汇总文件: ${summaryFilePath}`);
  } catch (error) {
    console.error("程序执行异常:", error.message);
  }
};

// 直接运行脚本时启动；保留导出以便验证并发池
if (require.main === module) {
  main();
}

module.exports = { runPagePool };
