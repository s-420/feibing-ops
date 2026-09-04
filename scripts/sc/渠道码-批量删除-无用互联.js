/**
 * @对象    渠道码
 * @动作    批量写
 * @风险    高
 * @来源    store-data-extractor/ShanChuWuYongQuDaoMaHuLian.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const fs = require("fs").promises;
const path = require("path");
const XLSX = require("xlsx");
const axios = require("axios");

// 配置信息 - 新增并发控制参数
const config = {
  excelPath: path.join(__dirname, "../input/互联删除无用渠道.xlsx"),
  logPath: path.join(__dirname, "../output/delete_channels_log.txt"),
  baseUrl: "https://vinci-api.feibing.tech/sc/v1/sellers/wwa9c5a585540b115b/places/",
  concurrency: 10, // 一次并发10条
  batchDelay: 1000, // 批次间延迟1秒（避免接口限流）
  headers: {
    accept: "*/*",
    "accept-language": "zh-CN,zh;q=0.9",
    authorization: "bearer __VINCI_TOKEN__",
    origin: "https://connect.feibing.tech",
    priority: "u=1, i",
    referer: "https://connect.feibing.tech/",
    "sec-ch-ua": '"Google Chrome";v="141", "Not?A_Brand";v="8", "Chromium";v="141"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"macOS"',
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-site",
    "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36",
  },
  cookies: {
    "x-token": "__VINCI_TOKEN__",
    acw_tc: "1a167a8a17615574197486702e436b3a3f1a3aef929b84345a1dd97e1cd6ef",
  },
};

// 记录日志信息
const logData = {
  startTime: new Date(),
  success: [],
  failed: [],
};

// 延迟函数，控制批次间隔
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// 读取Excel文件获取渠道ID（逻辑不变）
async function readChannelIds() {
  try {
    const workbook = XLSX.readFile(config.excelPath);
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    let idColumn = null;
    if (jsonData.length > 0) {
      const columns = Object.keys(jsonData[0]);
      idColumn = columns.find(
        (col) => col.trim().includes("渠道信息ID") || col.trim().includes("渠道ID")
      );
    }

    if (!idColumn) {
      throw new Error('未找到"渠道信息ID"列，请检查Excel文件');
    }

    return jsonData.map((item) => item[idColumn]).filter(Boolean);
  } catch (error) {
    console.error("读取Excel文件失败:", error.message);
    throw error;
  }
}

// 发送删除请求（逻辑不变）
async function deleteChannel(channelId) {
  try {
    const url = `${config.baseUrl}${channelId}`;
    const cookieStr = Object.entries(config.cookies)
      .map(([key, value]) => `${key}=${value}`)
      .join("; ");
    const headers = { ...config.headers, cookie: cookieStr };

    const response = await axios.delete(url, { headers });
    return { success: true, channelId, status: response.status };
  } catch (error) {
    return {
      success: false,
      channelId,
      error: error.message,
      status: error.response?.status || "未知",
    };
  }
}

// 批量删除渠道（修改为并发处理）
async function batchDeleteChannels(channelIds) {
  console.log(`开始处理${channelIds.length}个渠道，并发数: ${config.concurrency}`);
  
  // 将渠道ID分成多个批次（每批10条）
  const batches = [];
  for (let i = 0; i < channelIds.length; i += config.concurrency) {
    batches.push(channelIds.slice(i, i + config.concurrency));
  }

  // 逐批处理
  for (const [batchIndex, batch] of batches.entries()) {
    console.log(`\n处理第${batchIndex + 1}/${batches.length}批，共${batch.length}条渠道`);
    
    // 并发处理当前批次的所有渠道
    const batchResults = await Promise.all(
      batch.map((channelId) => deleteChannel(channelId))
    );

    // 记录批次结果
    batchResults.forEach((result) => {
      if (result.success) {
        logData.success.push({
          channelId: result.channelId,
          time: new Date(),
          status: result.status,
        });
        console.log(`渠道 ${result.channelId} 删除成功（状态码: ${result.status}）`);
      } else {
        logData.failed.push({
          channelId: result.channelId,
          time: new Date(),
          status: result.status,
          error: result.error,
        });
        console.log(`渠道 ${result.channelId} 删除失败（状态码: ${result.status}）: ${result.error}`);
      }
    });

    // 批次间延迟（最后一批不延迟）
    if (batchIndex < batches.length - 1) {
      console.log(`当前批次处理完成，等待${config.batchDelay}ms后处理下一批...`);
      await delay(config.batchDelay);
    }
  }
}

// 生成日志文件（逻辑不变）
async function generateLogFile() {
  logData.endTime = new Date();
  logData.total = logData.success.length + logData.failed.length;

  let logContent = `===== 渠道删除操作日志 =====\n`;
  logContent += `开始时间: ${logData.startTime.toLocaleString()}\n`;
  logContent += `结束时间: ${logData.endTime.toLocaleString()}\n`;
  logContent += `总处理数: ${logData.total}\n`;
  logContent += `成功数: ${logData.success.length}\n`;
  logContent += `失败数: ${logData.failed.length}\n\n`;

  logContent += `----- 成功删除的渠道 -----\n`;
  logData.success.forEach((item, index) => {
    logContent += `${index + 1}. 渠道ID: ${item.channelId} - 时间: ${item.time.toLocaleString()} - 状态码: ${item.status}\n`;
  });

  logContent += `\n----- 删除失败的渠道 -----\n`;
  logData.failed.forEach((item, index) => {
    logContent += `${index + 1}. 渠道ID: ${item.channelId} - 时间: ${item.time.toLocaleString()} - 状态码: ${item.status || "未知"} - 错误: ${item.error}\n`;
  });

  try {
    await fs.writeFile(config.logPath, logContent);
    console.log(`\n日志已生成: ${config.logPath}`);
  } catch (error) {
    console.error("生成日志文件失败:", error.message);
  }
}

// 主函数（逻辑不变）
async function main() {
  try {
    const channelIds = await readChannelIds();
    if (channelIds.length === 0) {
      console.log("未找到任何渠道ID，程序退出");
      return;
    }

    await batchDeleteChannels(channelIds);
    await generateLogFile();

    console.log("\n批量删除操作完成");
    console.log(`成功: ${logData.success.length} 个`);
    console.log(`失败: ${logData.failed.length} 个`);
  } catch (error) {
    console.error("程序执行出错:", error.message);
  }
}

// 启动程序
main();