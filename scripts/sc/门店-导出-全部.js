/**
 * @对象    门店
 * @动作    导出
 * @风险    低
 * @来源    store-data-extractor/DaoChuQuanBuMenDian.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const axios = require("axios");
const ExcelJS = require("exceljs");
const fs = require("fs-extra");
const path = require("path");

// 配置
const API_URL = "https://vinci-api.feibing.tech/sc/v1/sellers/wpV3cNDAAA7gXDSpALtCGXgFxSpYMBQQ/shops";
const PAGE_SIZE = 100; // 每次请求100条
const OUTPUT_DIR = path.join(__dirname, "../output");
const TEMP_DIR = path.join(OUTPUT_DIR, "门店数据临时表");

// 请求头
const headers = {
  accept: "*/*",
  "accept-language": "zh-CN,zh;q=0.9",
  "authorization": "bearer __VINCI_TOKEN__",
  origin: "https://connect.feibing.tech",
  priority: "u=1, i",
  referer: "https://connect.feibing.tech/",
  "sec-ch-ua": '"Not)A;Brand";v="8", "Chromium";v="138", "Google Chrome";v="138"',
  "sec-ch-ua-mobile": "?0",
  "sec-ch-ua-platform": '"macOS"',
  "sec-fetch-dest": "empty",
  "sec-fetch-mode": "cors",
  "sec-fetch-site": "same-site",
  "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36",
};

// Cookie
const cookie = "_clck=sj51u%7C2%7Cfxn%7C0%7C1974; x-token=__VINCI_TOKEN__";

// 初始化目录
async function initDirs() {
  await fs.ensureDir(OUTPUT_DIR);
  await fs.emptyDir(TEMP_DIR);
  await fs.ensureDir(TEMP_DIR);
}

// 发送请求获取数据
async function fetchData(pageNum) {
  try {
    const response = await axios.get(API_URL, {
      params: {
        pageSize: PAGE_SIZE,
        pageNum: pageNum,
        status: "",
        wxTagIds: "66b0c8f6b9fa2b0f92b711e8",
        // nameLike: "茶瀑布"
      },
      headers: {
        ...headers,
        cookie: cookie,
      },
      timeout: 15000
    });
    return response.data;
  } catch (error) {
    console.error(`获取第${pageNum}页数据失败，重试中...:`, error.message);
    await new Promise(resolve => setTimeout(resolve, 1000));
    return fetchData(pageNum);
  }
}

// 处理标签数据
function processTags(tags) {
  return {
    firstTag: tags?.[0]?.name || "",
    firstTagId: tags?.[0]?.id || "",
    secondTag: tags?.[1]?.name || "",
    secondTagId: tags?.[1]?.id || "",
    thirdTag: tags?.[2]?.name || "",
    thirdTagId: tags?.[2]?.id || ""
  };
}

// 处理单页数据
async function processPage(pageNum, data) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(`门店数据_${pageNum}`);

  // 表格列配置（含群ID）
  worksheet.columns = [
    { header: "ID", key: "id", width: 30 },
    { header: "名称", key: "name", width: 20 },
    { header: "省份", key: "province", width: 15 },
    { header: "城市", key: "city", width: 15 },
    { header: "区域", key: "area", width: 15 },
    { header: "地址", key: "address", width: 50 },
    { header: "状态", key: "status", width: 10 },
    { header: "一级战区", key: "firstTag", width: 40 },
    { header: "一级战区ID", key: "firstTagId", width: 30 },
    { header: "二级战区", key: "secondTag", width: 40 },
    { header: "二级战区ID", key: "secondTagId", width: 30 },
    { header: "三级战区", key: "thirdTag", width: 40 },
    { header: "三级战区ID", key: "thirdTagId", width: 30 },
    { header: "经度", key: "longitude", width: 20 },
    { header: "纬度", key: "latitude", width: 20 },
    { header: "绑定的客户群", key: "wxGroupChatName", width: 60 },
    { header: "群ID", key: "wxGroupChatId", width: 40 },
    { header: "关联员工", key: "wxStaffs", width: 80 },
    { header: "第三方门店code", key: "thirdShopCode", width: 20 },
    { header: "logo", key: "logo", width: 20 }
  ];

  // 处理数据并按客户群拆分
  data.forEach((shop) => {
    const baseData = {
      id: shop.id,
      name: shop.name,
      province: shop.location?.province?.name || "",
      city: shop.location?.city?.name || "",
      area: shop.location?.area?.name || "",
      address: shop.location?.address || "",
      status: shop.statusName || "",
      ...processTags(shop.tags),
      longitude: shop.location?.longitude || "",
      latitude: shop.location?.latitude || "",
      wxStaffs: shop.wxStaffs?.map(staff => staff.id).join(", ") || "",
      thirdShopCode: shop.metadata?.thirdShopCode || "",
      logo: shop.media[0]?.url || ""
    };

    // 拆分客户群数据
    const wxGroupChats = shop.wxGroupChats || [];
    if (wxGroupChats.length > 0) {
      wxGroupChats.forEach(chat => {
        worksheet.addRow({
          ...baseData,
          wxGroupChatName: chat.name || "",
          wxGroupChatId: chat.id || ""
        });
      });
    } else {
      worksheet.addRow({
        ...baseData,
        wxGroupChatName: "",
        wxGroupChatId: ""
      });
    }
  });

  const filePath = path.join(TEMP_DIR, `门店数据_${pageNum}.xlsx`);
  await workbook.xlsx.writeFile(filePath);
  console.log(`已生成第${pageNum}页数据，保存至: ${filePath}`);
  return filePath;
}

// 合并所有临时Excel文件
async function mergeExcelFiles(filePaths) {
  const mergedWorkbook = new ExcelJS.Workbook();
  const mergedWorksheet = mergedWorkbook.addWorksheet("所有门店数据汇总");

  // 汇总表列配置
  mergedWorksheet.columns = [
    { header: "ID", key: "id", width: 30 },
    { header: "名称", key: "name", width: 20 },
    { header: "省份", key: "province", width: 15 },
    { header: "城市", key: "city", width: 15 },
    { header: "区域", key: "area", width: 15 },
    { header: "地址", key: "address", width: 50 },
    { header: "状态", key: "status", width: 10 },
    { header: "一级战区", key: "firstTag", width: 40 },
    { header: "一级战区ID", key: "firstTagId", width: 30 },
    { header: "二级战区", key: "secondTag", width: 40 },
    { header: "二级战区ID", key: "secondTagId", width: 30 },
    { header: "三级战区", key: "thirdTag", width: 40 },
    { header: "三级战区ID", key: "thirdTagId", width: 30 },
    { header: "经度", key: "longitude", width: 20 },
    { header: "纬度", key: "latitude", width: 20 },
    { header: "绑定的客户群", key: "wxGroupChatName", width: 60 },
    { header: "群ID", key: "wxGroupChatId", width: 40 },
    { header: "关联员工", key: "wxStaffs", width: 80 },
    { header: "第三方门店code", key: "thirdShopCode", width: 20 },
    { header: "logo", key: "logo", width: 20 }
  ];

  // 合并数据
  for (const filePath of filePaths) {
    const workbook = await new ExcelJS.Workbook().xlsx.readFile(filePath);
    const worksheet = workbook.worksheets[0];

    for (let i = 2; i <= worksheet.rowCount; i++) {
      const row = worksheet.getRow(i);
      if (row.values[1]) {
        mergedWorksheet.addRow({
          id: row.values[1],
          name: row.values[2],
          province: row.values[3],
          city: row.values[4],
          area: row.values[5],
          address: row.values[6],
          status: row.values[7],
          firstTag: row.values[8],
          firstTagId: row.values[9],
          secondTag: row.values[10],
          secondTagId: row.values[11],
          thirdTag: row.values[12],
          thirdTagId: row.values[13],
          longitude: row.values[14],
          latitude: row.values[15],
          wxGroupChatName: row.values[16],
          wxGroupChatId: row.values[17],
          wxStaffs: row.values[18],
          thirdShopCode: row.values[19],
          logo: row.values[20]
        });
      }
    }
  }

  const mergedFilePath = path.join(OUTPUT_DIR, `门店数据汇总_${new Date().getTime()}.xlsx`);
  await mergedWorkbook.xlsx.writeFile(mergedFilePath);
  console.log(`所有数据已汇总，保存至: ${mergedFilePath}`);

  await fs.remove(TEMP_DIR);
  return mergedFilePath;
}

// 主函数
async function main() {
  try {
    console.log("开始获取门店数据...");
    await initDirs();

    const firstPageData = await fetchData(1);
    const total = firstPageData.total || 0;
    const totalPages = Math.ceil(total / PAGE_SIZE);

    console.log(`共发现${total}条数据，分为${totalPages}页`);

    const tempFiles = [];
    const firstPageFilePath = await processPage(1, firstPageData.data || []);
    tempFiles.push(firstPageFilePath);

    for (let pageNum = 2; pageNum <= totalPages; pageNum++) {
      console.log(`开始获取第${pageNum}页数据...`);
      const pageData = await fetchData(pageNum);
      const filePath = await processPage(pageNum, pageData.data || []);
      tempFiles.push(filePath);
    }

    await mergeExcelFiles(tempFiles);
    console.log("所有操作完成！");
  } catch (error) {
    console.error("操作失败:", error);
  }
}

main();