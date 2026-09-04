/**
 * @对象    桌台
 * @动作    批量写
 * @风险    高
 * @来源    store-data-extractor/PiLiangXiuGaiZuoTaiDuiYinKeFu.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

require("dotenv").config();
const xlsx = require('xlsx');
const axios = require('axios');
const path = require('path');
const fs = require('fs-extra');

/**
 * 桌台批量修改客服脚本（Windows / macOS / Linux 通用版）
 *
 * 改造说明（相对旧版）：
 * 旧版通过 execSync 拼接 curl 命令（bash 语法：单引号、反斜杠续行），
 * 在 Windows 的 cmd/PowerShell 下无法执行。
 * 新版改用 axios 直接发送 HTTP 请求，不再依赖系统 curl 和 shell 语法，
 * 三个平台行为完全一致，且避免了 shell 引号转义、进程创建开销等问题。
 *
 * 功能：从Excel读取门店ID和客服ID，通过API批量更新桌台对应的客服信息
 * 注意：请确保网络通畅且Excel文件格式正确，API返回code=0且message=成功时视为更新成功
 */

// ==================== 认证信息（过期后请整体替换） ====================
// 获取方式：浏览器登录 https://connect.feibing.tech 后，F12 打开开发者工具 ->
// Network -> 随便点一个 vinci-api 请求 -> Request Headers 里复制 authorization
// 中 "bearer " 后面的整串（不含 "bearer " 前缀）
// 优先读取项目根目录 .env 的凭据（推荐，全项目脚本共用一处，更新 token 无需改代码）
// .env 中 VINCI_AUTHORIZATION 填 authorization 整行值（含或不含 "bearer " 前缀均可）
// 未配置或为空时，回落到下方内置的历史 token（已过期）
const AUTH_TOKEN = (process.env.VINCI_AUTHORIZATION || `bearer __VINCI_TOKEN__`).replace(/^bearer\s+/i, '').trim();

// Cookie：同样从浏览器请求头里复制 cookie 的值（含 x-token）
const COOKIE = process.env.VINCI_COOKIE || '_clck=sj51u%7C2%7Cfxn%7C0%7C1974; x-token=__VINCI_TOKEN__';

// ==================== 配置项 ====================
const config = {
  excelPath: path.join(__dirname, '../input/桌台对应客服.xlsx'), // Excel文件路径（必填）
  stopAt: 0, // 控制修改到第几条停止（从1开始计数，0表示全部执行）
  batchUpdateType: 0, // 批量更新类型（默认值，具体含义请参考API文档）
  apiBaseUrl: 'https://vinci-api.feibing.tech/sc/v1/sellers/wwa9c5a585540b115b', // API基础地址
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
        console.error('   请重新登录 https://connect.feibing.tech，在项目根目录 .env 的 VINCI_AUTHORIZATION（和 VINCI_COOKIE）中更新凭据（全项目脚本共用）');
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
 * 调用API更新单个门店的桌台客服
 * @param {string} shopId 门店ID
 * @param {string} customerServiceId 客服ID
 * @returns {Promise<object>} API返回的JSON对象
 */
async function updateContact(shopId, customerServiceId) {
  const requestData = {
    batchUpdateType: config.batchUpdateType,
    shopIds: [shopId],
    customerService: JSON.stringify({ nickName: customerServiceId }),
    owners: [customerServiceId]
  };

  const response = await axios.post(
    `${config.apiBaseUrl}/places/${shopId}/batch/update/contact`,
    requestData,
    {
      headers: {
        'accept': 'application/json',
        'accept-language': 'zh-CN,zh;q=0.9',
        'authorization': `bearer ${AUTH_TOKEN}`,
        'content-type': 'application/json;charset=UTF-8',
        'origin': 'https://connect.feibing.tech',
        'referer': 'https://connect.feibing.tech/',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36',
        'cookie': COOKIE
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
 * 主函数：读取Excel数据并批量执行更新操作
 */
async function main() {
  // 初始化统计变量
  const stats = {
    total: 0,
    success: 0,
    failed: 0,
    skipped: 0,
    failedRecords: [] // 记录失败的行信息
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
    stats.total = data.length;
    console.log(`✅ 共读取到 ${stats.total} 条数据`);
    console.log(`📌 执行配置：将处理到第 ${config.stopAt || '全部'} 条数据`);
    console.log('------------------------------');

    // 4. 循环处理每条数据（按顺序执行）
    for (let i = 0; i < data.length; i++) {
      const rowNumber = i + 1; // 行号从1开始计数（与Excel显示一致）

      // 检查是否达到停止条件
      if (config.stopAt && rowNumber > config.stopAt) {
        console.log(`📌 已达到设定的停止行数 ${config.stopAt}，终止执行`);
        break;
      }

      // 获取当前行数据
      const row = data[i];
      const shopId = row['门店ID']?.toString().trim();
      const customerServiceId = row['客服ID']?.toString().trim();

      // 验证当前行数据完整性
      if (!shopId || !customerServiceId) {
        stats.skipped++;
        const reason = '门店ID或客服ID为空';
        stats.failedRecords.push({ rowNumber, shopId, customerServiceId, reason });
        console.log(`⚠️ 第 ${rowNumber} 行数据不完整（${reason}），已跳过`);
        continue;
      }

      console.log(`🔍 正在处理第 ${rowNumber} 行：`);
      console.log(`   门店ID: ${shopId}`);
      console.log(`   客服ID: ${customerServiceId}`);

      try {
        // 发送更新请求
        console.log('   正在发送更新请求...');
        const apiResponse = await updateContact(shopId, customerServiceId);

        // 判断是否成功
        if (apiResponse && apiResponse.code === 0 && apiResponse.message === '成功') {
          stats.success++;
          console.log(`✅ 第 ${rowNumber} 行更新成功`);
        } else {
          stats.failed++;
          // message过长时截断（避免服务端回显token刷屏）
          const rawMsg = apiResponse?.message?.toString() || '无返回';
          const msg = rawMsg.length > 80 ? rawMsg.substring(0, 80) + '...（已截断）' : rawMsg;
          const reason = `API返回异常（code=${apiResponse?.code}, message=${msg}）`;
          stats.failedRecords.push({ rowNumber, shopId, customerServiceId, reason });
          console.log(`❌ 第 ${rowNumber} 行更新失败（${reason}）`);
        }
      } catch (err) {
        stats.failed++;
        const reason = err.response
          ? `HTTP ${err.response.status} ${JSON.stringify(err.response.data).substring(0, 200)}`
          : (err.code ? `${err.code} ${err.message}` : err.message);
        stats.failedRecords.push({ rowNumber, shopId, customerServiceId, reason: reason.substring(0, 200) });
        console.error(`❌ 第 ${rowNumber} 行执行失败:`, reason);
      }

      // 请求间隔（最后一条不用等）
      if (i < data.length - 1 && !(config.stopAt && rowNumber >= config.stopAt)) {
        await sleep(config.requestIntervalMs);
      }

      console.log('------------------------------');
    }

    // 输出统计结果
    console.log('📊 执行结果统计:');
    console.log(`   总处理条数: ${stats.total}`);
    console.log(`   成功条数: ${stats.success}（${((stats.success / stats.total) * 100).toFixed(2)}%）`);
    console.log(`   失败条数: ${stats.failed}`);
    console.log(`   跳过条数: ${stats.skipped}`);

    // 打印所有失败的行记录
    if (stats.failedRecords.length > 0) {
      console.log('\n❌ 失败行详情:');
      stats.failedRecords.forEach(record => {
        console.log(`   行号: ${record.rowNumber}`);
        console.log(`   门店ID: ${record.shopId || '空'}`);
        console.log(`   客服ID: ${record.customerServiceId || '空'}`);
        console.log(`   失败原因: ${record.reason}`);
        console.log('   ------------------------------');
      });
    }

    console.log('🎉 脚本执行完成');
  } catch (err) {
    console.error('❌ 脚本执行出错:', err.message);
    console.log('请解决上述问题后重新运行脚本');
    process.exit(1);
  }
}

// 启动脚本
console.log('====== 桌台批量修改客服脚本启动（Windows/macOS/Linux 通用版） ======');
main();
