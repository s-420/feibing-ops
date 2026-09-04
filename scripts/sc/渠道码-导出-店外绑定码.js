/**
 * @对象    渠道码
 * @动作    导出
 * @风险    低
 * @来源    store-data-extractor/DaoChuDianWaiBangDingMa.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

require("dotenv").config();
const axios = require("axios");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

const config = {
  apiUrl:
    "https://vinci-api.feibing.tech/sc/v1/sellers/wwa9c5a585540b115b/places",
  pageSize: 100,
  concurrency: 10,
  startPage: 1,
  maxRetries: 3,
  retryDelay: 1000,
  requestTimeout: 30000,
  params: {
    types: "DESK,SINGLE,CHANNEL",
    catalogId: "69c280d94841606c242ca269",
    placeName: "店外-绑定码",
    sort: "modifiedTime,desc",
  },
  headers: {
    accept: "*/*",
    "accept-language": "zh-CN,zh;q=0.9",
    authorization:
      process.env.VINCI_AUTHORIZATION ||
      "bearer __VINCI_TOKEN__",
    cookie:
      process.env.VINCI_COOKIE ||
      "_clck=c2po6l%5E2%5Eg47%5E0%5E2259; x-token=__VINCI_TOKEN__",
    origin: "https://connect.feibing.tech",
    referer: "https://connect.feibing.tech/",
    "sec-ch-ua":
      '"Google Chrome";v="149", "Chromium";v="149", "Not)A;Brand";v="24"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"macOS"',
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-site",
    "user-agent":
      "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36",
  },
  outputDir: path.join(__dirname, "../output/店外绑定码"),
  columns: [
    { header: "码名称", key: "name", width: 22 },
    { header: "码ID", key: "id", width: 28 },
    { header: "绑定时间", key: "boundTime", width: 20 },
    { header: "门店ID", key: "shopId", width: 28 },
    { header: "绑定门店", key: "shopName", width: 30 },
    { header: "添加时间", key: "createdTime", width: 20 },
  ],
};

const delay = (milliseconds) =>
  new Promise((resolve) => setTimeout(resolve, milliseconds));

const ensureDirExists = (dirPath) => {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
};

const parseTime = (timestamp) => {
  if (!timestamp) return null;
  // Excel 日期本身不携带时区；转换为北京时间对应的表格日期值
  const date = new Date(Number(timestamp) + 8 * 60 * 60 * 1000);
  date.setUTCMilliseconds(0);
  return Number.isNaN(date.getTime()) ? null : date;
};

const mapPlaceToRow = (item) => ({
  name: item.name || "",
  id: item.id || "",
  boundTime: parseTime(item.modifiedTime),
  shopId: item.shop?.id || "",
  shopName: item.shop?.name || "",
  createdTime: parseTime(item.createdTime),
});

const createExcelFile = async (rows, filePath) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("店外绑定码");
  worksheet.columns = config.columns;
  worksheet.addRows(rows);
  const headerRow = worksheet.getRow(1);
  headerRow.height = 24;
  headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
  headerRow.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF1F4E78" },
  };
  headerRow.alignment = { vertical: "middle", horizontal: "center" };
  worksheet.getColumn("boundTime").numFmt = "yyyy-mm-dd hh:mm:ss";
  worksheet.getColumn("createdTime").numFmt = "yyyy-mm-dd hh:mm:ss";
  worksheet.views = [{ state: "frozen", ySplit: 1 }];
  worksheet.autoFilter = {
    from: "A1",
    to: `F${Math.max(1, rows.length + 1)}`,
  };
  await workbook.xlsx.writeFile(filePath);
};

const fetchPage = async (pageNum) => {
  let lastError;

  for (let attempt = 1; attempt <= config.maxRetries; attempt += 1) {
    try {
      const response = await axios.get(config.apiUrl, {
        params: {
          pageNum,
          pageSize: config.pageSize,
          ...config.params,
        },
        headers: config.headers,
        timeout: config.requestTimeout,
      });

      if (response.data.code !== 0) {
        throw new Error(
          `接口返回错误 ${response.data.code}: ${response.data.message || "未知错误"}`
        );
      }

      return {
        total: Number(response.data.total) || 0,
        data: Array.isArray(response.data.data) ? response.data.data : [],
      };
    } catch (error) {
      lastError = error;
      console.warn(
        `第 ${pageNum} 页第 ${attempt}/${config.maxRetries} 次请求失败: ${error.message}`
      );
      if (attempt < config.maxRetries) {
        await delay(config.retryDelay * attempt);
      }
    }
  }

  throw lastError;
};

// 固定并发池：一个任务完成后，立即领取下一页
const runPagePool = async (pages, concurrency, worker) => {
  let nextIndex = 0;

  const runWorker = async () => {
    while (nextIndex < pages.length) {
      const currentIndex = nextIndex;
      nextIndex += 1;
      await worker(pages[currentIndex]);
    }
  };

  await Promise.all(
    Array.from(
      { length: Math.min(concurrency, pages.length) },
      () => runWorker()
    )
  );
};

const main = async () => {
  ensureDirExists(config.outputDir);

  console.log("正在获取第 1 页并确认总页数...");
  const firstPage = await fetchPage(1);
  const totalPages = Math.ceil(firstPage.total / config.pageSize);
  console.log(
    `共 ${firstPage.total} 条，${totalPages} 页，每页 ${config.pageSize} 条，并发数 ${config.concurrency}`
  );

  if (totalPages === 0) {
    console.log("接口没有返回数据，程序结束");
    return;
  }

  const rowsByPage = new Array(totalPages + 1);
  const failedPages = [];
  const savePage = async (pageNum, data) => {
    const rows = data.map(mapPlaceToRow);
    rowsByPage[pageNum] = rows;
    const filePath = path.join(
      config.outputDir,
      `店外绑定码_第${pageNum}页.xlsx`
    );
    await createExcelFile(rows, filePath);
    console.log(`第 ${pageNum}/${totalPages} 页已保存，共 ${rows.length} 条`);
  };

  if (config.startPage <= 1) {
    await savePage(1, firstPage.data);
  }

  const loopStartPage = Math.max(2, config.startPage);
  const pages = Array.from(
    { length: Math.max(0, totalPages - loopStartPage + 1) },
    (_, index) => loopStartPage + index
  );

  await runPagePool(pages, config.concurrency, async (pageNum) => {
    try {
      console.log(`开始处理第 ${pageNum} 页...`);
      const result = await fetchPage(pageNum);
      await savePage(pageNum, result.data);
    } catch (error) {
      failedPages.push(pageNum);
      console.error(`第 ${pageNum} 页最终失败: ${error.message}`);
    }
  });

  if (failedPages.length > 0) {
    throw new Error(
      `以下分页导出失败，未生成汇总文件: ${failedPages.sort((a, b) => a - b).join(", ")}`
    );
  }

  const allRows = rowsByPage.slice(config.startPage).flat();
  const summaryFilePath = path.join(
    config.outputDir,
    "店外绑定码_汇总.xlsx"
  );
  await createExcelFile(allRows, summaryFilePath);

  console.log("\n导出完成！");
  console.log(`分页文件目录: ${config.outputDir}`);
  console.log(`汇总文件: ${summaryFilePath}`);
  console.log(`汇总记录数: ${allRows.length}`);
};

if (require.main === module) {
  main().catch((error) => {
    console.error("程序执行失败:", error.message);
    process.exitCode = 1;
  });
}

module.exports = {
  config,
  mapPlaceToRow,
  createExcelFile,
  runPagePool,
};
