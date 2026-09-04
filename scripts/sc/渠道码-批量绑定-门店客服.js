/**
 * @对象    渠道码
 * @动作    批量写
 * @风险    高
 * @来源    store-data-extractor/ZhuMaBangDingMenDianKeFu.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

/** 按“门店客服批量刷”映射逐码绑定客服手机号。 */
require('dotenv').config();

const fs = require('fs');
const path = require('path');
const axios = require('axios');
const { readMappings, config: mappingConfig } = require('./PiLiangShuaXinMenDianKeFuYiDuiYi');

const outputDir = path.join(__dirname, '../output');
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const stamp = () => new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14);
const parseContact = (value) => { try { return JSON.parse(value || '{}'); } catch (_) { return null; } };

function client() {
  const auth = String(process.env.VINCI_AUTHORIZATION || '').trim();
  if (!auth) throw new Error('缺少 VINCI_AUTHORIZATION');
  return axios.create({
    headers: {
      authorization: /^bearer\s+/i.test(auth) ? auth : `bearer ${auth}`,
      'content-type': 'application/json;charset=UTF-8',
      origin: 'https://connect.feibing.tech',
      referer: 'https://connect.feibing.tech/',
    },
    timeout: 60000,
    validateStatus: () => true,
  });
}

function assertSuccess(response, action) {
  if (response.status < 200 || response.status >= 300 || response.data?.code !== 0) {
    throw new Error(`${action}: HTTP ${response.status}, ${response.data?.message || `code=${response.data?.code}`}`);
  }
  return response.data;
}

async function listPlaces(api, shopId) {
  const response = await api.get(`https://vinci-api.feibing.tech/sc/v1/sellers/${mappingConfig.sellerId}/places`, {
    params: { current: 1, pageSize: 100, pageNum: 1, shopId, catalogId: '', types: 'DESK,SINGLE,CHANNEL,GROUP' },
  });
  return assertSuccess(response, `查询门店 ${shopId}`).data || [];
}

async function updatePlace(api, placeId, contact) {
  const response = await api.put(
    `https://vinci-api.feibing.tech/sc/v1/sellers/${mappingConfig.sellerId}/places/${placeId}/metadata/contact`,
    contact,
  );
  assertSuccess(response, `更新码 ${placeId}`);
}

async function main() {
  if (process.env.CUSTOMER_SERVICE_WRITE_CONFIRM !== '逐码绑定193家门店客服') throw new Error('缺少逐码写入二次确认');
  const mappings = readMappings();
  const api = client();
  fs.mkdirSync(outputDir, { recursive: true });
  const runStamp = stamp();
  const backupPath = path.join(outputDir, `备份_逐码绑定门店客服_${runStamp}.jsonl`);
  const resultPath = path.join(outputDir, `结果_逐码绑定门店客服_${runStamp}.jsonl`);
  let total = 0;
  let skipped = 0;
  let updated = 0;
  let failed = 0;

  for (let storeIndex = 0; storeIndex < mappings.length; storeIndex += 1) {
    const mapping = mappings[storeIndex];
    const places = await listPlaces(api, mapping.shopId);
    total += places.length;
    for (const place of places) {
      const contact = parseContact(place.metadata?.contact);
      if (!contact) {
        failed += 1;
        fs.appendFileSync(resultPath, `${JSON.stringify({ storeName: mapping.storeName, placeId: place.id, status: 'failed', reason: 'contact 无法解析' })}\n`);
        continue;
      }
      const correct = Array.isArray(contact.owners) && contact.owners.length === 1 && String(contact.owners[0]) === mapping.customerAccount;
      if (correct) { skipped += 1; continue; }
      const oldOwners = Array.isArray(contact.owners) ? [...contact.owners] : [];
      fs.appendFileSync(backupPath, `${JSON.stringify({ storeName: mapping.storeName, shopId: mapping.shopId, placeId: place.id, oldOwners, oldContact: contact, targetPhone: mapping.customerAccount })}\n`);
      contact.owners = [mapping.customerAccount];
      try {
        await updatePlace(api, place.id, contact);
        updated += 1;
        fs.appendFileSync(resultPath, `${JSON.stringify({ storeName: mapping.storeName, placeId: place.id, status: 'updated', oldOwners, targetPhone: mapping.customerAccount })}\n`);
      } catch (error) {
        failed += 1;
        fs.appendFileSync(resultPath, `${JSON.stringify({ storeName: mapping.storeName, placeId: place.id, status: 'failed', reason: error.message })}\n`);
      }
      await delay(250);
    }
    console.log(`门店进度 ${storeIndex + 1}/${mappings.length}，累计码 ${total}，已更新 ${updated}，原本正确 ${skipped}，失败 ${failed}`);
    await delay(150);
  }
  console.log(JSON.stringify({ total, updated, skipped, failed, backupPath, resultPath }));
  if (failed > 0) process.exitCode = 2;
}

main().catch((error) => { console.error(`执行失败: ${error.message}`); process.exit(1); });
