/**
 * @对象    桌台
 * @动作    导出
 * @风险    低
 * @来源    store-data-extractor/DaoChuZhiDinMenDianZuoTai.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const fsExtra = require('fs-extra');
const axios = require('axios');

// 导出指定门店对应的门店号渠道信息
const config = {
  apiUrl: 'https://vinci-api.feibing.tech/sc/v1/sellers/wpV3cNDAAA7gXDSpALtCGXgFxSpYMBQQ/places',
  headers: {
    'accept': '*/*',
    'accept-language': 'zh-CN,zh;q=0.9',
    "authorization": "bearer __VINCI_TOKEN__",
    'Origin': 'https://connect.feibing.tech',
    'Referer': 'https://connect.feibing.tech/',
    'Sec-Fetch-Dest': 'empty',
    'Sec-Fetch-Mode': 'cors',
    'Sec-Fetch-Site': 'same-site',
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36',
    'sec-ch-ua': '"Not)A;Brand";v="8", "Chromium";v="138", "Google Chrome";v="138"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"macOS"'
  },
  cookies: '_clck=sj51u%7C2%7Cfxn%7C0%7C1974; x-token=__VINCI_TOKEN__',
  pageSize: 100,
  maxRetries: 2,
  concurrency: 5,
  tempDir: path.join(__dirname, '../temp_data_two'),
  outputDir: path.join(__dirname, '../output'),
  outputFile: '指定门店桌台信息.xlsx',
  inputFile: path.join(__dirname, '../input/指定门店.xlsx'),
  shopIdColumnName: '店铺id'
};

// 确保目录存在
async function ensureDirs() {
  await fsExtra.ensureDir(path.dirname(config.inputFile));
  await fsExtra.emptyDir(config.tempDir);
  await fsExtra.ensureDir(config.tempDir);
  await fsExtra.ensureDir(config.outputDir);
}

// 记录错误日志（仅打印关键错误）
function logError(message) {
  const timestamp = new Date().toISOString();
  const errorLogPath = path.join(config.outputDir, 'error_log.txt');
  fs.appendFileSync(errorLogPath, `[${timestamp}] ${message}\n`);
  console.error(`❌ 错误：${message}`);
}

// 读取输入表格中的shopId列
function readShopIdsFromExcel() {
  try {
    if (!fs.existsSync(config.inputFile)) {
      throw new Error(`输入文件不存在：${config.inputFile}`);
    }

    const workbook = XLSX.readFile(config.inputFile);
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    const shopIds = [...new Set(
      jsonData
        .map(item => item[config.shopIdColumnName]?.toString().trim())
        .filter(id => id && id !== 'undefined' && id !== 'null')
    )];

    if (shopIds.length === 0) {
      throw new Error(`未从${config.shopIdColumnName}列中读取到有效数据`);
    }

    console.log(`📋 成功读取${shopIds.length}个有效店铺ID`);
    return shopIds;
  } catch (error) {
    logError(`读取店铺ID失败：${error.message}`);
    throw error;
  }
}

// 调整接口调用：支持传入shopId参数（去除重试过程打印）
async function fetchPage(pageNum, shopId) {
  for (let retry = 1; retry <= config.maxRetries; retry++) {
    try {
      const headers = { ...config.headers };
      headers.Cookie = config.cookies;

      const response = await axios.get(config.apiUrl, {
        params: {
          current: pageNum,
          pageSize: config.pageSize,
          pageNum: pageNum,
          shopId: shopId,
          catalogId: '',
          types: 'DESK,SINGLE,CHANNEL,GROUP'
          // CHANNEL
        },
        headers: headers,
        timeout: 10000
      });

      if (response.data.code !== 0) {
        throw new Error(`接口错误：${response.data.message || '未知错误'}`);
      }

      return {
        data: response.data.data || [],
        total: response.data.total || 0
      };
    } catch (error) {
      const errorMsg = error.response?.data?.message || error.message;
      if (retry < config.maxRetries) {
        await new Promise(resolve => setTimeout(resolve, 1000));
      } else {
        logError(`店铺ID=${shopId}第${pageNum}页请求失败：${errorMsg}`);
        return null;
      }
    }
  }
}

// 处理数据（保留shopId关联）
function processData(items, shopId) {
  return items
    .filter(item => item.name && /^门店号/.test(item.name))
    .map(item => ({
      // '目标shopId': shopId,
      '门店id': item.shop?.id || '',
      '门店名称': item.shop?.name || '',
      'id': item.id || '',
      'name': item.name || '',
      'link字段': (() => {
        try {
          if (item.metadata?.contact) {
            const contact = JSON.parse(item.metadata.contact);
            if (contact.welMsg?.attachments?.length) {
              const linkItem = contact.welMsg.attachments.find(a => a.type === 'link');
              return linkItem?.link || '';
            }
          }
          return '';
        } catch (e) {
          return '';
        }
      })(),
      // "福利官ID": item.metadata?.staffWxUserid || '',
      "福利官ID": (() => {
        if (item.metadata?.contact) {
          const contact = JSON.parse(item.metadata.contact);
          if (contact.owners.length) {
            return contact.owners[0] || '';
          }
        }
        return '';
      })(),
      "logo": item.metadata?.qrCodeAvatar || '',
      '自定义标签': (() => {
        try {
          if (item.metadata?.contactDerived) {
            const contactDerived = JSON.parse(item.metadata.contactDerived);
            // if (contact.welMsg?.attachments?.length) {
            //   const linkItem = contact.welMsg.attachments.find(a => a.type === 'link');
            //   return linkItem?.link || '';
            // }
            if (contactDerived.wxCpTags?.length > 0) {
              // contactDerived.wxCpTags是否包含制定标签id
              const tagId = '668f4d28b9fa2b05d92a1834';
              const tag = contactDerived.wxCpTags.find(t => t.id === tagId);
              if (tag) {
                return tag.name;
              }
              return '';
            }
          }
          return '无自定义标签';
        } catch (e) {
          return '无自定义标签';
        }
      })(),
      // "关联客服开关": (() => {
      //   try {
      //     return JSON.parse(item.metadata?.contact)?.preFilter?.openFilter ? '开启' : '关闭';
      //   } catch (e) {
      //     return '';
      //   }
      // })()
    }));
}

// 保存单页数据（去除保存成功的冗余日志）
function savePageData(data, pageNum, shopId) {
  if (!data.length) return true;

  const filePath = path.join(config.tempDir, `shop_${shopId}_page_${pageNum}.xlsx`);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(data), '数据');
  XLSX.writeFile(workbook, filePath);
  return true;
}

async function mergeAllData() {
  try {
    if (!fs.existsSync(config.tempDir)) {
      throw new Error(`临时目录不存在：${config.tempDir}`);
    }

    // 筛选符合格式的文件（shop_xxx_page_xxx.xlsx，其中xxx可以是字母、数字、下划线等）
    const files = fs.readdirSync(config.tempDir)
      .filter(f => /^shop_[a-zA-Z0-9]+_page_\d+\.xlsx$/.test(f)) // 调整正则匹配字母数字组合
      .sort((a, b) => {
        // 提取shopId（支持字母数字组合）
        const aShopMatch = a.match(/shop_([a-zA-Z0-9]+)_page_\d+\.xlsx/);
        const bShopMatch = b.match(/shop_([a-zA-Z0-9]+)_page_\d+\.xlsx/);
        const aShop = aShopMatch ? aShopMatch[1] : '';
        const bShop = bShopMatch ? bShopMatch[1] : '';

        // 先按shopId字符串排序
        if (aShop !== bShop) {
          return aShop.localeCompare(bShop);
        }

        // 再按页码排序
        const aPageMatch = a.match(/shop_[a-zA-Z0-9]+_page_(\d+)\.xlsx/);
        const bPageMatch = b.match(/shop_[a-zA-Z0-9]+_page_(\d+)\.xlsx/);
        const aPage = aPageMatch ? parseInt(aPageMatch[1], 10) : 0;
        const bPage = bPageMatch ? parseInt(bPageMatch[1], 10) : 0;

        return aPage - bPage;
      });

    if (!files.length) {
      console.log(`ℹ️ 无符合条件的数据可合并`);
      return;
    }

    const allData = files.flatMap(file => {
      const fileFullPath = path.join(config.tempDir, file);
      try {
        return XLSX.utils.sheet_to_json(XLSX.readFile(fileFullPath).Sheets['数据']);
      } catch (e) {
        logError(`合并文件${file}失败：${e.message}`);
        return [];
      }
    });

    const outputPath = path.join(config.outputDir, config.outputFile);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(allData), '指定门店数据');
    XLSX.writeFile(workbook, outputPath);

    console.log(`🎉 合并完成，共${allData.length}条数据，结果已保存至：${outputPath}`);
  } catch (error) {
    logError(`合并数据失败：${error.message}`);
    throw error;
  }
}


// 获取单个shopId的总页数（去除冗余日志）
async function getTotalPagesByShopId(shopId) {
  try {
    const firstPage = await fetchPage(1, shopId);
    return firstPage ? Math.ceil(firstPage.total / config.pageSize) : 1;
  } catch (error) {
    logError(`获取店铺ID=${shopId}总页数失败：${error.message}`);
    return 0;
  }
}

// 处理单个shopId的所有分页数据（仅保留关键状态）
async function processSingleShop(shopId) {
  try {
    const totalPages = await getTotalPagesByShopId(shopId);
    if (totalPages === 0) {
      console.log(`ℹ️ 店铺ID=${shopId}：无数据，已跳过`);
      return;
    }

    for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
      const pageData = await fetchPage(pageNum, shopId);
      if (pageData) {
        const processedData = processData(pageData.data, shopId);
        savePageData(processedData, pageNum, shopId);
      }
    }
    console.log(`✅ 店铺ID=${shopId}：处理完成（共${totalPages}页）`);
  } catch (error) {
    logError(`处理店铺ID=${shopId}时出错：${error.message}`);
    console.log(`⚠️ 店铺ID=${shopId}：处理失败，已跳过`);
  }
}

// 分批处理shopId（仅保留批次统计）
async function processShopsInBatches(shopIds) {
  const batchSize = config.concurrency;
  const totalBatches = Math.ceil(shopIds.length / batchSize);

  console.log(`📦 共${shopIds.length}个店铺ID，将分${totalBatches}批处理（每批${batchSize}个）`);

  for (let batch = 0; batch < totalBatches; batch++) {
    const start = batch * batchSize;
    const end = start + batchSize;
    const currentBatch = shopIds.slice(start, end);

    console.log(`\n🚀 开始处理第${batch + 1}/${totalBatches}批`);
    await Promise.allSettled(
      currentBatch.map(shopId => processSingleShop(shopId))
    );
    console.log(`✅ 第${batch + 1}/${totalBatches}批处理完成`);
  }
}

// 主函数：精简打印
async function main() {
  try {
    await ensureDirs();
    console.log(`🚀 开始处理指定门店数据（并发数：${config.concurrency}）`);

    const shopIds = readShopIdsFromExcel();
    if (!shopIds.length) {
      console.log(`ℹ️ 无有效店铺ID，终止流程`);
      return;
    }

    await processShopsInBatches(shopIds);
    console.log(`\n🔗 开始合并所有数据...`);
    await mergeAllData();

    console.log(`\n✅ 全部流程完成！`);
  } catch (error) {
    console.error(`💥 主流程中断：${error.message}`);
  } finally {
    console.log(`\n📝 所有操作已执行完毕`);
  }
}

// 执行主函数
main().catch(error => {
  logError(`主函数未捕获错误：${error.message}`);
  process.exit(1);
});