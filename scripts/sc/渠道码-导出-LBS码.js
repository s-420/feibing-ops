/**
 * @对象    渠道码
 * @动作    导出
 * @风险    低
 * @来源    store-data-extractor/DaoChuQuanBuLBSMa.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const axios = require("axios");
const ExcelJS = require("exceljs");
const fs = require("fs-extra");
const path = require("path");

// 配置参数
const config = {
  // API配置
  apiUrl:
    "https://vinci-api.feibing.tech/sc/v1/sellers/wwa9c5a585540b115b/places",
  pageSize: 100, // 每页100条数据
  maxPagesPerBatch: 100, // 每批次处理100页

  // 店铺ID来源Excel
  shopIdExcelPath: path.join(__dirname, "../input/一点点全部门店.xlsx"),
  shopIdColumnName: "店铺id", // Excel中店铺ID列的表头名称

  // 请求头和Cookie
  headers: {
    accept: "*/*",
    "accept-language": "zh-CN,zh;q=0.9",
    authorization:
      "bearer __VINCI_TOKEN__",
    origin: "https://connect.feibing.tech",
    priority: "u=1, i",
    referer: "https://connect.feibing.tech/",
    "sec-ch-ua":
      '"Not)A;Brand";v="8", "Chromium";v="138", "Google Chrome";v="138"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"macOS"',
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-site",
    "user-agent":
      "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36",
  },
  cookie:
    "_clck=sj51u%7C2%7Cfxn%7C0%7C1974; x-token=__VINCI_TOKEN__",

  // 文件配置
  outputDir: path.join(__dirname, "../output/渠道数据汇总"),
  tempDir: path.join(__dirname, "../output/渠道数据临时表"),
  finalFileName: `渠道数据汇总_${new Date().getTime()}.xlsx`,
  sheetName: "渠道数据",

  // 表格列配置（按需求定义）
  columns: [
    { header: "门店ID", key: "shopId", width: 30 },
    { header: "门店名称", key: "shopName", width: 25 },
    { header: "渠道码ID", key: "channelID", width: 20 },
    { header: "渠道信息ID", key: "channelInfoID", width: 20 },
    { header: "渠道名称", key: "channelName", width: 20 },
    { header: "对应客服ID", key: "staffId", width: 40 },
    { header: "入群链接", key: "link", width: 80 },
  ],
};

// 初始化目录
async function initDirs() {
  await fs.ensureDir(config.outputDir);
  // 只确保临时目录存在，不清除已有文件（保留历史临时文件）
  await fs.ensureDir(config.tempDir);
}

// 读取Excel中的店铺ID
async function readShopIds() {
  try {
    console.log(`读取店铺ID列表: ${config.shopIdExcelPath}`);

    if (!(await fs.pathExists(config.shopIdExcelPath))) {
      throw new Error(`店铺ID文件不存在: ${config.shopIdExcelPath}`);
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(config.shopIdExcelPath);

    // 读取第一个工作表
    const worksheet = workbook.worksheets[0];
    if (!worksheet) {
      throw new Error("Excel文件中没有工作表");
    }

    // 查找店铺ID列的索引
    let shopIdColumnIndex = -1;
    worksheet.getRow(1).eachCell((cell, index) => {
      if (cell.value === config.shopIdColumnName) {
        shopIdColumnIndex = index;
      }
    });

    if (shopIdColumnIndex === -1) {
      throw new Error(`未找到"${config.shopIdColumnName}"列`);
    }

    // 读取所有店铺ID
    const shopIds = [];
    for (let i = 2; i <= worksheet.rowCount; i++) {
      const cell = worksheet.getCell(i, shopIdColumnIndex);
      if (cell.value) {
        shopIds.push(cell.value.toString().trim());
      }
    }

    console.log(`成功读取${shopIds.length}个店铺ID`);
    return shopIds;
  } catch (error) {
    console.error("读取店铺ID失败:", error.message);
    throw error;
  }
}

// 发送请求获取数据
async function fetchData(pageNum, shopId) {
  try {
    console.log(`📡 请求店铺[${shopId}]第${pageNum}页数据...`);
    const response = await axios.get(config.apiUrl, {
      params: {
        current: pageNum,
        pageSize: config.pageSize,
        pageNum: pageNum,
        shopId: shopId, // 使用动态店铺ID
        catalogId: "",
        // type: "DESK,SINGLE,CHANNEL"
      },
      headers: {
        ...config.headers,
        cookie: config.cookie,
      },
    });

    if (response.data.code !== 0) {
      throw new Error(`接口返回错误: ${response.data.message || "未知错误"}`);
    }

    console.log(
      `📥 店铺[${shopId}]第${pageNum}页请求成功，返回${
        response.data.data?.length || 0
      }条数据`
    );
    return {
      data: response.data.data || [],
      total: response.data.total || 0,
    };
  } catch (error) {
    console.error(`❌ 店铺[${shopId}]第${pageNum}页请求失败:`, error.message);
    // 重试机制
    await new Promise((resolve) => setTimeout(resolve, 1000));
    return fetchData(pageNum, shopId);
  }
}

// 处理单个店铺的单页数据并生成临时Excel
async function processShopPage(shopId, pageNum, data, globalCounter) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(`渠道数据_${pageNum}`);

  // 设置列
  worksheet.columns = config.columns;

  // 处理数据并更新计数器
  data.forEach((item, index) => {
    let link = ""
    try {
      link = "无配置链接";

      if (item.metadata?.contact) {
        const contact = JSON.parse(item.metadata.contact);

        if (contact?.welMsg) {
          if (contact.welMsg.attachments?.length) {
            const linkItem = contact.welMsg.attachments.find(
              (a) => a.type === "link"
            );
            if (linkItem?.link) {
              link = linkItem.link;
            }
          }
        }
      }
    } catch (e) {
      // 解析失败（如JSON格式错误）时，也返回“无链接”
      link = "无链接";
      console.warn(`⚠️ ID=${item.id} 解析失败：${e.message}`);
    }

    // 计算当前是全局第几条数据
    const currentCount = globalCounter + index + 1;
    worksheet.addRow({
      shopId: item.shop?.id || shopId, // 使用传入的shopId作为默认值
      shopName: item.shop?.name || "",
      channelInfoID: item.id || "",
      channelName: item.name || "",
      channelID: item.metadata?.contact_id || "",
      staffId: item.metadata?.staffWxUserid || "",
      link: link,
    });
    // 每10条打印一次（避免日志过多），最后一条强制打印
    if (currentCount % 10 === 0 || index === data.length - 1) {
      console.log(`📊 当前执行到第${currentCount}条数据（店铺[${shopId}]）`);
    }
  });

  // 保存临时文件，直接使用门店ID作为文件名前缀
  const filePath = path.join(
    config.tempDir,
    `${shopId}_渠道数据_${pageNum}.xlsx`
  );
  await workbook.xlsx.writeFile(filePath);
  console.log(
    `💾 店铺[${shopId}]第${pageNum}页数据已保存到临时文件: ${filePath}`
  );

  // 返回本页处理的数据量
  return data.length;
}

// 处理单个店铺的所有数据
async function processShop(shopId, shopIndex, totalShops) {
  try {
    console.log(
      `\n===== 开始处理店铺(${shopIndex}/${totalShops}): ${shopId} =====`
    );

    // 获取第一页数据，确定总页数
    const firstPageResult = await fetchData(1, shopId);
    const total = firstPageResult.total;
    const totalPages = Math.ceil(total / config.pageSize);

    console.log(`店铺[${shopId}]共发现${total}条数据，分为${totalPages}页`);

    // 初始化本店铺的数据计数器
    let shopCounter = 0;
    const tempFiles = [];

    // 处理第一页
    const firstPageCount = await processShopPage(
      shopId,
      1,
      firstPageResult.data,
      shopCounter
    );
    tempFiles.push(path.join(config.tempDir, `${shopId}_渠道数据_1.xlsx`));
    shopCounter += firstPageCount;

    // 处理剩余页面（每批次100页）
    let currentPage = 2;
    while (currentPage <= totalPages) {
      const batchEndPage = Math.min(
        currentPage + config.maxPagesPerBatch - 1,
        totalPages
      );
      console.log(
        `开始处理店铺[${shopId}]批次：第${currentPage}页至第${batchEndPage}页`
      );

      for (let pageNum = currentPage; pageNum <= batchEndPage; pageNum++) {
        const pageResult = await fetchData(pageNum, shopId);
        const pageCount = await processShopPage(
          shopId,
          pageNum,
          pageResult.data,
          shopCounter
        );
        tempFiles.push(
          path.join(config.tempDir, `${shopId}_渠道数据_${pageNum}.xlsx`)
        );
        shopCounter += pageCount;
      }

      currentPage = batchEndPage + 1;
    }

    console.log(`店铺[${shopId}]处理完成，共处理${shopCounter}条数据`);
    return { tempFiles, count: shopCounter };
  } catch (error) {
    console.error(`处理店铺[${shopId}]时出错:`, error.message);
    return { tempFiles: [], count: 0 };
  }
}

// 合并所有临时文件（修复link字段映射错误）
async function mergeExcelFiles(allTempFiles) {
  console.log("\n开始汇总所有店铺数据...");

  const mergedWorkbook = new ExcelJS.Workbook();
  const mergedWorksheet = mergedWorkbook.addWorksheet(config.sheetName);

  // 设置汇总表列
  mergedWorksheet.columns = config.columns;

  // 合并数据并计数
  let totalRows = 0;
  for (const filePath of allTempFiles) {
    try {
      const workbook = await new ExcelJS.Workbook().xlsx.readFile(filePath);
      const worksheet = workbook.worksheets[0];

      // 从第二行开始读取（跳过表头）
      for (let i = 2; i <= worksheet.rowCount; i++) {
        const row = worksheet.getRow(i);
        if (row.values[1]) {
          // 确保行有数据
          mergedWorksheet.addRow({
            shopId: row.values[1],
            shopName: row.values[2],
            channelID: row.values[3],
            channelInfoID: row.values[4],
            channelName: row.values[5],
            staffId: row.values[6],
            link: row.values[7], // 修复link字段映射，对应第5列
          });
          totalRows++;
        }
      }
      console.log(`合并完成: ${path.basename(filePath)}（累计${totalRows}条）`);
    } catch (error) {
      console.error(`合并文件${filePath}失败:`, error.message);
    }
  }

  // 保存汇总文件
  const mergedFilePath = path.join(config.outputDir, config.finalFileName);
  await mergedWorkbook.xlsx.writeFile(mergedFilePath);
  console.log(
    `🎉 所有店铺数据汇总完成，共${totalRows}条记录，保存至: ${mergedFilePath}`
  );

  // 移除临时文件删除操作，保留临时文件
  console.log(`💾 临时文件保留在: ${config.tempDir}`);
  return mergedFilePath;
}

// 主函数
async function main() {
  try {
    console.log("===== 开始获取多店铺渠道数据 =====");
    await initDirs();

    // 读取所有店铺ID
    const shopIds = await readShopIds();
    if (shopIds.length === 0) {
      console.log("没有找到任何店铺ID，程序退出");
      return;
    }

    // 处理所有店铺并累计总数据量
    const allTempFiles = [];
    let globalTotalCount = 0;

    for (let i = 0; i < shopIds.length; i++) {
      const shopId = shopIds[i];
      const { tempFiles, count } = await processShop(
        shopId,
        i + 1,
        shopIds.length
      );
      allTempFiles.push(...tempFiles);
      globalTotalCount += count;
      console.log(
        `===== 店铺(${i + 1}/${
          shopIds.length
        })处理完成，累计处理${globalTotalCount}条数据 =====\n`
      );
    }

    // 合并所有文件
    await mergeExcelFiles(allTempFiles);
    console.log("\n✅ 所有操作完成！");
  } catch (error) {
    console.error("\n💥 操作失败:", error.message);
    // 出错时也不删除临时文件
    console.log(`⚠️ 临时文件保留在: ${config.tempDir}`);
  }
}

// 执行主函数
main();
