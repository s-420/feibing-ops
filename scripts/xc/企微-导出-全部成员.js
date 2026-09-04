/**
 * @对象    企微
 * @动作    导出
 * @风险    低
 * @来源    store-data-extractor/DaoChuQuanBuChengYuan.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const axios = require('axios');

// 确保output文件夹及子文件夹存在
const outputDir = path.join(__dirname, '../output');
const splitDir = path.join(outputDir, 'split-files');
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}
if (!fs.existsSync(splitDir)) {
  fs.mkdirSync(splitDir, { recursive: true });
}

// 读取Excel文件
const readExcel = (filePath) => {
  try {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
  } catch (error) {
    console.error('读取Excel文件失败:', error);
    process.exit(1);
  }
};

// 写入Excel文件
const writeExcel = (data, outputPath, sheetName = '结果') => {
  try {
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    XLSX.writeFile(workbook, outputPath);
    console.log(`结果已保存至: ${outputPath}`);
  } catch (error) {
    console.error('写入Excel文件失败:', error);
  }
};

// 调用API查询
const queryContact = async (phoneNumber) => {
  console.log(`查询手机号: ${phoneNumber}`);
  try {
    const response = await axios.post(
      'https://vinci-api.feibing.tech/xc/v1/cps/token/contact/search',
      {
        queryWord: phoneNumber,
        sellerId: 'wpV3cNDAAA7gXDSpALtCGXgFxSpYMBQQ'
      },
      {
        headers: {
          'accept': 'application/json',
          'accept-language': 'zh-CN,zh;q=0.9',
          'authorization': 'bearer __VINCI_TOKEN__',
          'content-type': 'application/json;charset=UTF-8',
          'origin': 'https://connect.feibing.tech',
          'referer': 'https://connect.feibing.tech/',
          'sec-ch-ua': '"Not)A;Brand";v="8", "Chromium";v="138", "Google Chrome";v="138"',
          'sec-ch-ua-mobile': '?0',
          'sec-ch-ua-platform': '"macOS"',
          'sec-fetch-dest': 'empty',
          'sec-fetch-mode': 'cors',
          'sec-fetch-site': 'same-site',
          'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36'
        },
        withCredentials: true
      }
    );
    return response.data;
  } catch (error) {
    console.error(`查询手机号 ${phoneNumber} 时出错:`, error.message);
    return null;
  }
};

const processContact = async (contact) => {
  const { name, phone, department } = contact;

  if (!phone) {
    console.log(`跳过没有手机号的记录: ${name}`);
    return {
      '姓名': name,
      '手机号': '无手机号',
      '部门': department || '无部门信息',
      '用户ID': '无手机号'
    };
  }

  console.log(`正在查询 ${name}（${department || '未知部门'}，${phone}）...`);
  const response = await queryContact(phone);

  let userId = '未找到';
  if (response && response.code === 0 && response.data?.user?.userid?.length > 0) {
    userId = response.data.user.userid[0];
  } else {
    console.log(`查询失败: ${response?.message || '未知错误'}`);
  }

  console.log(`完成 ${name} 的查询，用户ID: ${userId}`);
  return {
    '姓名': name,
    '手机号': phone,
    '部门': department || '无部门信息',
    '用户ID': userId
  };
};

// 固定并发池：任意一条完成后立即补充下一条，并保持结果顺序不变
const runWithConcurrency = async (items, concurrency, worker, onCompleted) => {
  const results = new Array(items.length);
  let nextIndex = 0;

  const runWorker = async () => {
    while (nextIndex < items.length) {
      const currentIndex = nextIndex;
      nextIndex += 1;
      results[currentIndex] = await worker(items[currentIndex], currentIndex);
      if (onCompleted) {
        onCompleted(currentIndex, results);
      }
    }
  };

  await Promise.all(
    Array.from(
      { length: Math.min(concurrency, items.length) },
      () => runWorker()
    )
  );

  return results;
};

// 主函数
const main = async () => {
  const contactFilePath = path.join(__dirname, '../input/私域项目组通讯录.xlsx');
  const contacts = readExcel(contactFilePath);
  const concurrency = 10;
  const splitSize = 100;

  console.log(`开始处理，共 ${contacts.length} 条记录，并发数: ${concurrency}`);

  const normalizedContacts = contacts.map(contact => ({
    name: contact['姓名'],
    phone: contact['手机'],
    department: contact['部门']
  }));

  const completed = new Array(normalizedContacts.length).fill(false);
  let nextSplitStart = 0;
  let splitFileCount = 0;

  const exportReadySplits = (completedIndex, results) => {
    completed[completedIndex] = true;

    // 仅当连续的100条全部完成时才导出，确保拆分文件保持通讯录原顺序
    while (nextSplitStart + splitSize <= results.length) {
      const splitEnd = nextSplitStart + splitSize;
      const isSplitCompleted = completed
        .slice(nextSplitStart, splitEnd)
        .every(Boolean);

      if (!isSplitCompleted) break;

      splitFileCount += 1;
      const splitData = results.slice(nextSplitStart, splitEnd);
      const splitFilePath = path.join(
        splitDir,
        `查询结果_${splitFileCount}.xlsx`
      );
      writeExcel(splitData, splitFilePath, `结果_${splitFileCount}`);
      nextSplitStart = splitEnd;
    }
  };

  const allResults = await runWithConcurrency(
    normalizedContacts,
    concurrency,
    processContact,
    exportReadySplits
  );

  // 导出最后不足100条的剩余数据
  if (nextSplitStart < allResults.length) {
    splitFileCount += 1;
    const splitData = allResults.slice(nextSplitStart);
    const splitFilePath = path.join(splitDir, `查询结果_${splitFileCount}.xlsx`);
    writeExcel(splitData, splitFilePath, `结果_${splitFileCount}`);
  }

  const summaryPath = path.join(outputDir, '查询结果汇总.xlsx');
  writeExcel(allResults, summaryPath, '全部结果');

  console.log(`所有处理完成！共生成 ${splitFileCount} 个拆分文件和 1 个汇总文件`);
  console.log(`拆分文件目录: ${splitDir}`);
  console.log(`汇总文件路径: ${summaryPath}`);
};

if (require.main === module) {
  main().catch(error => console.error('程序执行出错:', error));
}

module.exports = { runWithConcurrency };
