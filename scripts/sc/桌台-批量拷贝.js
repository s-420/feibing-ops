/**
 * @对象    桌台
 * @动作    批量写
 * @风险    高
 * @来源    store-data-extractor/PiLiangKaoBeiZuoTai.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

require("dotenv").config();
const xlsx = require('xlsx');
const axios = require('axios');
const path = require('path');
const fs = require('fs-extra');

/**
 * 批量拷贝桌台脚本（Windows / macOS / Linux 通用版）
 *
 * 改造说明（相对旧版）：
 * 旧版通过 execSync 拼接 curl 命令（bash 语法），Windows 的 cmd/PowerShell 无法执行。
 * 新版改用 axios 直接发送 HTTP 请求，不再依赖系统 curl 和 shell 语法，跨平台通用。
 *
 * 配置项
 */

// ==================== 认证信息（过期后请整体替换） ====================
// 获取方式：浏览器登录管理后台后，F12 打开开发者工具 ->
// Network -> 随便点一个 vinci-api 请求 -> Request Headers 里复制 authorization
// 中 "bearer " 后面的整串（不含 "bearer " 前缀）
// 优先读取项目根目录 .env 的凭据（推荐，全项目脚本共用一处，更新 token 无需改代码）
// .env 中 VINCI_AUTHORIZATION 填 authorization 整行值（含或不含 "bearer " 前缀均可）
// 未配置或为空时，回落到下方内置的历史 token（已过期）
const AUTH_TOKEN = (process.env.VINCI_AUTHORIZATION || `bearer __VINCI_TOKEN__`).replace(/^bearer\s+/i, '').trim();

// ==================== 配置项 ====================
const config = {
  excelPath: path.join(__dirname, '../input/拷贝桌台.xlsx'), // Excel文件路径
  // 固定的源店铺ID
  sourceShopId: "6309e821c3242c25f986b812",
  placeNames: ["校园店招新"],
  // 接口配置
  apiUrl: 'https://vinci-api.feibing.tech/sc/v1/sellers/wwa9c5a585540b115b/shops/copy/place/source/to/target',
  useShopStaffId: true, // 固定参数
  requestIntervalMs: 300, // 每条请求之间的间隔（毫秒），避免请求过快触发限流
  timeoutMs: 30000, // 单条请求超时时间（毫秒）
  ignoreTokenExpiryWarning: false, // token过期时是否仍然强制执行（默认false：检测到过期直接终止）
};

/**
 * 检查 AUTH_TOKEN 是否过期（仅支持 JWT 格式 token）
 * 过期则打印警告；ignoreTokenExpiryWarning=false 时直接终止脚本
 */
function checkTokenExpiry() {
  try {
    const parts = AUTH_TOKEN.split('.');
    if (parts.length !== 3) return; // 不是标准 JWT 格式，跳过检查
    // 兼容 base64url：补齐 padding 并替换字符
    const payloadStr = Buffer.from(parts[1].replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString('utf8');
    const payload = JSON.parse(payloadStr);
    if (payload && payload.exp) {
      const expDate = new Date(payload.exp * 1000);
      if (Date.now() > payload.exp * 1000) {
        console.error(`❌ Token 已过期！过期时间: ${expDate.toLocaleString('zh-CN')}`);
        console.error('   请重新登录管理后台，在项目根目录 .env 的 VINCI_AUTHORIZATION 中更新凭据（全项目脚本共用）');
        if (!config.ignoreTokenExpiryWarning) {
          process.exit(1);
        }
        console.error('   （已按配置忽略过期警告，继续执行，但请求大概率全部失败）');
      } else {
        console.log(`✅ Token 有效期至: ${expDate.toLocaleString('zh-CN')}`);
      }
    }
  } catch (e) {
    // token 解析失败不影响执行，仅跳过检查
  }
}

/**
 * 调用API拷贝桌台
 * @param {object} requestData 请求数据
 * @returns {Promise<object>} API返回的JSON对象
 */
async function copyPlace(requestData) {
  const response = await axios.post(
    config.apiUrl,
    requestData,
    {
      headers: {
        'accept': 'application/json',
        'accept-language': 'zh-CN,zh;q=0.9',
        'authorization': `bearer ${AUTH_TOKEN}`,
        'content-type': 'application/json;charset=UTF-8',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36'
      },
      timeout: config.timeoutMs,
      // 与旧版 curl 行为保持一致：HTTP 状态码非2xx时不抛异常，把响应体交给上层按 code 字段判断
      validateStatus: () => true
    }
  );

  return response.data;
}

/**
 * 简单休眠工具函数
 * @param {number} ms 休眠毫秒数
 */
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * 主函数：读取Excel数据并批量执行拷贝操作
 */
async function main() {
  // 初始化统计变量
  const stats = {
    total: 0,          // 总操作次数（行 × placeNames数量）
    totalRows: 0,      // 总行数
    success: 0,
    failed: 0,
    skipped: 0,
    failedRecords: []  // 记录失败详情
  };

  try {
    // 0. 检查token是否过期
    checkTokenExpiry();

    // 1. 检查Excel文件是否存在
    console.log(`开始检查文件: ${config.excelPath}`);
    if (!await fs.pathExists(config.excelPath)) {
      throw new Error(`❌ Excel文件不存在，请确认路径是否正确: ${config.excelPath}`);
    }
    console.log('✅ 文件检查通过');

    // 2. 读取并解析Excel文件
    console.log('开始读取Excel数据...');
    const workbook = xlsx.readFile(config.excelPath);
    const sheetName = workbook.SheetNames[0]; // 取第一个工作表
    console.log(`使用工作表: ${sheetName}`);

    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet); // 转换为JSON数组

    // 3. 数据有效性检查
    if (!data.length) {
      throw new Error('❌ Excel文件中没有有效数据，请检查表格内容');
    }
    stats.totalRows = data.length;
    stats.total = stats.totalRows * config.placeNames.length;
    console.log(`✅ 共读取到 ${stats.totalRows} 行数据，将执行 ${stats.total} 次拷贝操作`);
    console.log(`📌 使用固定源店铺ID: ${config.sourceShopId}`);
    console.log(`📌 将自动使用的placeName: ${config.placeNames.join('、')}`);
    console.log('------------------------------');

    // 4. 循环处理每条数据
    for (let i = 0; i < data.length; i++) {
      const rowNumber = i + 1; // 行号从1开始计数
      const row = data[i];

      // 从Excel提取必要字段
      const targetShopsStr = row['targetShops']?.toString().trim();
      const staffUserId = row['staffUserId']?.toString().trim(); // 读取客服ID

      // 验证必要字段完整性
      const missingFields = [];
      if (!targetShopsStr) missingFields.push('targetShops');
      if (!staffUserId) missingFields.push('staffUserId');

      if (missingFields.length > 0) {
        stats.skipped += config.placeNames.length; // 一行对应的所有操作都跳过
        const reason = `缺少必要字段: ${missingFields.join(', ')}`;
        stats.failedRecords.push({
          rowNumber,
          reason,
          targetShops: targetShopsStr,
          staffUserId,
          placeNames: config.placeNames // 记录本次跳过的所有placeName
        });
        console.log(`⚠️ 第 ${rowNumber} 行数据不完整（${reason}），已跳过该行所有 ${config.placeNames.length} 次操作`);
        console.log('------------------------------');
        continue;
      }

      // 处理targetShops（解析为数组格式）
      let targetShopsIds;
      try {
        // 支持两种格式：直接逗号分隔或JSON数组
        if (targetShopsStr.startsWith('[')) {
          targetShopsIds = JSON.parse(targetShopsStr);
        } else {
          targetShopsIds = targetShopsStr.split(',').map(id => id.trim());
        }

        if (!Array.isArray(targetShopsIds) || targetShopsIds.length === 0) {
          throw new Error('目标店铺ID解析后不是有效数组');
        }
      } catch (parseErr) {
        stats.skipped += config.placeNames.length; // 一行对应的所有操作都跳过
        const reason = `targetShops格式错误：${parseErr.message}`;
        stats.failedRecords.push({
          rowNumber,
          reason,
          targetShops: targetShopsStr,
          staffUserId,
          placeNames: config.placeNames
        });
        console.log(`⚠️ 第 ${rowNumber} 行${reason}，已跳过该行所有 ${config.placeNames.length} 次操作`);
        console.log('------------------------------');
        continue;
      }

      console.log(`🔍 开始处理第 ${rowNumber} 行（目标店铺: [${targetShopsIds.join(', ')}]，客服ID: ${staffUserId}）`);

      // 对当前行的目标店铺，分别使用每个placeName执行拷贝操作
      for (const placeName of config.placeNames) {
        console.log(`   正在处理 ${placeName} ...`);

        // 构造请求数据（补充staffUserId）
        const requestData = {
          sourceShopId: config.sourceShopId,
          placeName,
          targetShops: { ids: targetShopsIds.join(', ') },
          staffUserId: staffUserId, // 从Excel读取的客服ID
          // useShopStaffId: config.useShopStaffId // 恢复固定参数
        };

        try {
          // 发送请求
          console.log('   正在发送请求...');
          const apiResponse = await copyPlace(requestData);

          // 判断是否成功
          if (apiResponse && apiResponse.code === 0 && apiResponse.message === '成功') {
            stats.success++;
            console.log(`✅ ${placeName} 执行成功`);
          } else {
            stats.failed++;
            // message过长时截断（避免服务端回显token刷屏）
            const rawMsg = apiResponse?.message?.toString() || '无返回';
            const msg = rawMsg.length > 80 ? rawMsg.substring(0, 80) + '...（已截断）' : rawMsg;
            const reason = `API返回异常（code=${apiResponse?.code}, message=${msg}）`;
            stats.failedRecords.push({
              rowNumber,
              reason,
              placeName,
              targetShops: targetShopsIds,
              staffUserId
            });
            console.log(`❌ ${placeName} 执行失败（${reason}）`);
          }
        } catch (err) {
          stats.failed++;
          const reason = err.response
            ? `HTTP ${err.response.status} ${JSON.stringify(err.response.data).substring(0, 200)}`
            : (err.code ? `${err.code} ${err.message}` : err.message);
          stats.failedRecords.push({
            rowNumber,
            reason: reason.substring(0, 200),
            placeName,
            targetShops: targetShopsIds,
            staffUserId
          });
          console.error(`❌ ${placeName} 执行失败:`, reason.substring(0, 200));
        }

        // 请求间隔（该行最后一个操作后不用等）
        await sleep(config.requestIntervalMs);
      }

      console.log(`✅ 第 ${rowNumber} 行的 ${config.placeNames.length} 次操作已处理完成`);
      console.log('------------------------------');
    }

    // 输出统计结果
    console.log('📊 执行结果统计:');
    console.log(`   总行数: ${stats.totalRows}`);
    console.log(`   总操作次数: ${stats.total}`);
    console.log(`   成功次数: ${stats.success}（${((stats.success/stats.total)*100).toFixed(2)}%）`);
    console.log(`   失败次数: ${stats.failed}`);
    console.log(`   跳过次数: ${stats.skipped}`);

    // 打印失败详情
    if (stats.failedRecords.length > 0) {
      console.log('\n❌ 失败详情:');
      stats.failedRecords.forEach(record => {
        console.log(`   行号: ${record.rowNumber}`);
        console.log(`   placeName: ${record.placeName || 'N/A'}`);
        console.log(`   客服ID: ${record.staffUserId}`);
        console.log(`   目标店铺ID: [${Array.isArray(record.targetShops) ? record.targetShops.join(', ') : record.targetShops}]`);
        console.log(`   失败原因: ${record.reason}`);
        console.log('   ------------------------------');
      });
    }

    console.log('🎉 脚本执行完成');
  } catch (err) {
    console.error('❌ 脚本执行出错:', err.message);
    process.exit(1);
  }
}

// 启动脚本
console.log('====== 批量拷贝桌台脚本启动（Windows/macOS/Linux 通用版） ======');
main();
