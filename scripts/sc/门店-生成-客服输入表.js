/**
 * @对象    门店
 * @动作    导出
 * @风险    低
 * @来源    store-data-extractor/ShengChengKeFuShuRuBiao.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

/**
 * 生成「客服号替换」脚本的输入表
 *
 * 用途：客户需求表里没有飞冰店铺ID，只有「三方ID」。本脚本读后台导出的总表，
 *       按「三方ID」匹配出店铺ID，再取客户表的「客服联系电话（企微id）」列，
 *       生成 TiHuanMenDianKeFuHao.js 需要的三列表格（门店名称 / 店铺ID / 客服手机号）。
 *
 * 纯本地 Excel 操作，不调用 API，无需 token。
 * 匹配键优先级：三方ID > 三方门店ID。
 */

const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// ⚠ 路径请用正斜杠 `/` 写（Windows 也支持），不要用反斜杠 `\`——
//   JS 字符串里 `\d`、`\202` 会被当成转义符，导致路径被破坏成 D:download... 这种错误。
const config = {
  // 总表（茶瀑布后台导出的全部门店，含「店铺id」+「三方ID」）
  totalTablePath: 'D:/download/changelink/20260828_554de7c47ddf4e07b8c1e9d7e3b008b4.xlsx',
  totalShopIdCol: '店铺id',
  totalThirdIdCol: '三方ID',
  totalThirdShopIdCol: '三方门店ID',

  // 客户需求表（含「三方id」+「门店名」+「客服联系电话（企微id）」）
  clientTablePath: 'D:/download/changelink/更换链接.xlsx',
  clientThirdIdCol: '三方id',
  clientThirdShopIdCol: '三方门店ID',
  clientNameCol: '门店名',
  clientPhoneCol: '客服联系电话（企微id）',

  // 输出（客服替换脚本的输入表，列名必须精确为「门店名称 / 店铺ID / 客服手机号」）
  outputPath: path.join(__dirname, '../input', '门店ID和客服手机号_茶瀑布.xlsx'),
  // 未匹配门店清单
  unmatchedPath: path.join(__dirname, '../output', '未匹配门店_客服.xlsx'),
};

function readRows(filePath) {
  const wb = xlsx.readFile(filePath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (rows.length === 0) throw new Error(`文件为空: ${filePath}`);
  return rows;
}

function colIndex(header, name) {
  const idx = header.map(h => String(h).trim()).indexOf(name);
  if (idx === -1) throw new Error(`找不到列「${name}」，实际表头: ${header.join(' | ')}`);
  return idx;
}

function norm(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}

function main() {
  // 1. 读总表，建立 三方ID / 三方门店ID -> 店铺id 映射
  const trows = readRows(config.totalTablePath);
  const thead = trows[0];
  const tShopId = colIndex(thead, config.totalShopIdCol);
  const tThirdId = colIndex(thead, config.totalThirdIdCol);
  const tThirdShopId = colIndex(thead, config.totalThirdShopIdCol);

  const byThirdId = new Map();
  const byThirdShopId = new Map();
  for (const r of trows.slice(1)) {
    const shopId = norm(r[tShopId]);
    const thirdId = norm(r[tThirdId]);
    const thirdShopId = norm(r[tThirdShopId]);
    if (!shopId) continue;
    if (thirdId && !byThirdId.has(thirdId)) byThirdId.set(thirdId, shopId);
    if (thirdShopId && !byThirdShopId.has(thirdShopId)) byThirdShopId.set(thirdShopId, shopId);
  }
  console.log(`总表: ${trows.length - 1} 行；三方ID 映射 ${byThirdId.size} 条，三方门店ID 映射 ${byThirdShopId.size} 条`);

  // 2. 读客户表，逐行匹配
  const crows = readRows(config.clientTablePath);
  const chead = crows[0];
  const cThirdId = colIndex(chead, config.clientThirdIdCol);
  const cThirdShopId = colIndex(chead, config.clientThirdShopIdCol);
  const cName = colIndex(chead, config.clientNameCol);
  const cPhone = colIndex(chead, config.clientPhoneCol);

  const matched = [];
  const unmatched = [];
  for (const r of crows.slice(1)) {
    const name = norm(r[cName]);
    const thirdId = norm(r[cThirdId]);
    const thirdShopId = norm(r[cThirdShopId]);
    const phone = norm(r[cPhone]);
    if (!name) continue; // 跳过空行

    let shopId = '';
    let via = '';
    if (thirdId && byThirdId.has(thirdId)) {
      shopId = byThirdId.get(thirdId);
      via = '三方ID';
    } else if (thirdShopId && byThirdShopId.has(thirdShopId)) {
      shopId = byThirdShopId.get(thirdShopId);
      via = '三方门店ID';
    }

    if (shopId && phone) {
      matched.push({ name, shopId, phone, via });
    } else if (shopId && !phone) {
      unmatched.push({ name, thirdId, thirdShopId, reason: '缺少客服手机号' });
    } else {
      unmatched.push({ name, thirdId, thirdShopId, reason: '未匹配到店铺ID' });
    }
  }

  console.log(`客户表匹配: 成功 ${matched.length} 个，未匹配 ${unmatched.length} 个`);

  // 3. 输出输入表（若旧文件存在，先备份）
  if (fs.existsSync(config.outputPath)) {
    const ts = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14);
    const backupPath = config.outputPath.replace(/\.xlsx$/i, `_备份_${ts}.xlsx`);
    fs.copyFileSync(config.outputPath, backupPath);
    console.log(`已备份旧输入表到: ${backupPath}`);
  }

  const outRows = matched.map(m => ({ '门店名称': m.name, '店铺ID': m.shopId, '客服手机号': m.phone }));
  const outWs = xlsx.utils.json_to_sheet(outRows);
  const outWb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(outWb, outWs, 'Sheet1');

  try {
    xlsx.writeFile(outWb, config.outputPath);
    console.log(`已生成输入表: ${config.outputPath}（${outRows.length} 行）`);
  } catch (e) {
    if (e.code === 'EBUSY' || e.code === 'EACCES') {
      const ts = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14);
      const altPath = config.outputPath.replace(/\.xlsx$/i, `_${ts}.xlsx`);
      xlsx.writeFile(outWb, altPath);
      console.log(`⚠ 目标文件被 Excel 占用，已改写到: ${altPath}`);
      console.log('  （关闭 Excel 后可手动改名为目标文件名，或直接使用此文件）');
    } else {
      throw e;
    }
  }

  // 4. 未匹配清单（若有）
  if (unmatched.length > 0) {
    const uRows = unmatched.map(u => ({ '门店名': u.name, '三方id': u.thirdId, '三方门店ID': u.thirdShopId, '原因': u.reason }));
    const uWs = xlsx.utils.json_to_sheet(uRows);
    const uWb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(uWb, uWs, 'Sheet1');
    xlsx.writeFile(uWb, config.unmatchedPath);
    console.log(`未匹配门店清单已保存: ${config.unmatchedPath}`);
    console.log('--- 未匹配门店 ---');
    unmatched.forEach(u => console.log(`  ${u.name} | ${u.reason}`));
  }
}

main();
