/**
 * @对象    企微
 * @动作    批量写
 * @风险    高
 * @来源    store-data-extractor/PiLiangJiHuoZhangHu.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

/**
 * 批量激活账户（茶瀑布）
 *
 * 业务：给茶瀑布门店的「客服账号」（企微账号）批量激活。激活的输入是客服手机号，
 *       系统通过手机号查到员工（企微 userID），再消耗一个激活码完成激活。
 *
 * 完整流程（每个客服手机号）：
 *   ① 手机号 → 企微 userID：POST /xc/v1/cps/token/contact/search { queryWord: 手机号, sellerId }
 *   ② 获取激活码：           GET  /sc/v1/sellers/{sellerId}/licenses/get/next/active/code → data = "LA200001..."
 *   ③ 激活账户：             POST /sc/v1/sellers/{sellerId}/licenses/active/account
 *                           body = { userid: 企微userID, corpid: sellerId, active_code: 激活码 }
 *
 * 成功判定：code===0 && data.ok===true
 *
 * 兜底说明：
 *   - 激活会真实消耗激活码（license），不可逆，故不提供回滚。
 *   - 提供 dryRun 预演：只做「手机号→userID」映射（不拿激活码、不激活），
 *     先把哪些能映射、哪些映射失败列出来，人工确认后再真写。
 *   - 详细日志：逐客服记录 手机号 → userID → 激活码 → 激活结果。
 *
 * 输入：客户需求表（含「门店名」+「客服联系电话（企微id）」两列）
 */

require('dotenv').config();

const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const axios = require('axios');

const config = {
  sellerId: 'wpV3cNDAAA3ZRQDQZt0UsR9wORkx0thg', // 茶瀑布 corpid
  // 客户需求表路径（含「门店名」+「客服联系电话（企微id）」列）—— 用正斜杠 /
  clientTablePath: 'D:/download/changelink/更换链接.xlsx',
  clientNameCol: '门店名',
  clientPhoneCol: '客服联系电话（企微id）',
  logPath: path.join(__dirname, '../output', '批量激活账户_茶瀑布_执行日志.txt'),
  // ★★ 预演开关：true = 只做手机号→userID 映射，不拿激活码、不激活
  dryRun: false,
  requestDelayMs: 300, // 每次请求间隔，避免限流
  headers: {
    'accept': '*/*',
    'accept-language': 'zh-CN,zh;q=0.9',
    'authorization': process.env.VINCI_AUTHORIZATION || '',
    'origin': 'https://connect.feibing.tech',
    'referer': 'https://connect.feibing.tech/',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36',
  },
};

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function logToFile(message) {
  try {
    if (!fs.existsSync(path.dirname(config.logPath))) {
      fs.mkdirSync(path.dirname(config.logPath), { recursive: true });
    }
    const timestamp = new Date().toLocaleString('zh-CN', {
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false
    });
    fs.appendFileSync(config.logPath, `[${timestamp}] ${message}\n`, 'utf8');
  } catch (error) {
    console.error('写入日志失败:', error.message);
  }
}

// ① 手机号 → 企微 userID
async function phoneToUserid(phone) {
  const response = await axios.post(
    'https://vinci-api.feibing.tech/xc/v1/cps/token/contact/search',
    { queryWord: phone, sellerId: config.sellerId },
    {
      headers: { ...config.headers, 'content-type': 'application/json;charset=UTF-8' },
      timeout: 20000,
    }
  );
  if (response.data.code !== 0) {
    throw new Error(`手机号 ${phone} 查询员工失败: ${response.data.message}`);
  }
  const userid = response.data?.data?.user?.userid?.[0];
  if (!userid) {
    throw new Error(`手机号 ${phone} 未找到对应员工（未进企业？）`);
  }
  return userid;
}

// ② 获取下一个激活码
async function getActiveCode() {
  const response = await axios.get(
    `https://vinci-api.feibing.tech/sc/v1/sellers/${config.sellerId}/licenses/get/next/active/code`,
    { headers: config.headers, timeout: 20000 }
  );
  if (response.data.code !== 0) {
    throw new Error(`获取激活码失败: ${response.data.message}`);
  }
  return response.data.data; // "LA200001..."
}

// ③ 激活账户
async function activeAccount(userid, activeCode) {
  const response = await axios.post(
    `https://vinci-api.feibing.tech/sc/v1/sellers/${config.sellerId}/licenses/active/account`,
    { userid, corpid: config.sellerId, active_code: activeCode },
    {
      headers: { ...config.headers, 'content-type': 'application/json;charset=UTF-8' },
      timeout: 20000,
    }
  );
  return response.data;
}

// 处理单个客服
async function processRow(row) {
  const name = String(row[config.clientNameCol] ?? '').trim();
  const phone = String(row[config.clientPhoneCol] ?? '').trim();

  if (!phone) {
    logToFile(`跳过：门店「${name || '未知'}」缺少客服手机号`);
    console.log(`跳过: ${name || '未知'} (无手机号)`);
    return 'skipped';
  }

  logToFile(`\n开始处理: ${name || ''} (手机号 ${phone})`);
  console.log(`处理中: ${name || phone}`);

  // ① 手机号 → userID（预演和真写都做，这一步是安全的）
  let userid;
  try {
    userid = await phoneToUserid(phone);
  } catch (e) {
    logToFile(`门店「${name}」- 映射失败: ${e.message}`);
    console.log(`已完成: ${name} (映射失败)`);
    return 'failed';
  }
  logToFile(`门店「${name}」- 手机号 ${phone} → 企微 userID: ${userid}`);

  // 预演模式：只到这里，不拿激活码、不激活
  if (config.dryRun) {
    logToFile(`门店「${name}」- 【预演】将激活 ${userid}（未拿激活码、未激活）`);
    console.log(`已完成: ${name} (预演)`);
    return 'dryrun';
  }

  // ② 获取激活码
  let activeCode;
  try {
    activeCode = await getActiveCode();
  } catch (e) {
    logToFile(`门店「${name}」- 获取激活码失败: ${e.message}`);
    console.log(`已完成: ${name} (拿激活码失败)`);
    return 'failed';
  }
  logToFile(`门店「${name}」- 获取到激活码: ${activeCode}`);

  // ③ 激活
  try {
    const result = await activeAccount(userid, activeCode);
    if (result && result.code === 0 && result.data && result.data.ok === true) {
      logToFile(`门店「${name}」- ✅ 激活成功: ${userid} (激活码 ${activeCode})`);
      console.log(`已完成: ${name} (成功)`);
      return 'success';
    }
    const msg = result?.message || result?.data?.errmsg || '未知原因';
    logToFile(`门店「${name}」- ❌ 激活失败: ${msg}`);
    console.log(`已完成: ${name} (失败)`);
    return 'failed';
  } catch (e) {
    logToFile(`门店「${name}」- 激活请求失败: ${e.message}`);
    console.log(`已完成: ${name} (失败)`);
    return 'failed';
  }
}

async function main() {
  const summary = { success: 0, failed: 0, skipped: 0, dryrun: 0 };

  try {
    logToFile('===== 程序开始执行 =====');
    console.log(`程序开始执行...（茶瀑布，预演: ${config.dryRun ? '开' : '关'}）`);

    const workbook = xlsx.readFile(config.clientTablePath);
    const rows = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
    logToFile(`成功读取 ${rows.length} 条数据`);
    console.log(`共读取到 ${rows.length} 条数据，开始处理...`);

    for (const row of rows) {
      try {
        const result = await processRow(row);
        summary[result] = (summary[result] || 0) + 1;
      } catch (e) {
        logToFile(`处理失败: ${e.message}`);
        summary.failed++;
      }
      if (config.requestDelayMs > 0) {
        await delay(config.requestDelayMs);
      }
    }

    logToFile('===== 程序执行结束 =====');
    logToFile(`汇总: 成功 ${summary.success} / 失败 ${summary.failed} / 跳过 ${summary.skipped} / 预演 ${summary.dryrun}`);
    console.log(`\n程序执行完成，详细日志已保存到: ${config.logPath}`);
    console.log(`汇总: 成功 ${summary.success} / 失败 ${summary.failed} / 跳过 ${summary.skipped} / 预演 ${summary.dryrun}`);
  } catch (e) {
    const msg = `程序执行出错: ${e.message}`;
    logToFile(msg);
    console.error(msg);
    process.exit(1);
  }
}

main();
