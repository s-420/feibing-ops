/**
 * @对象    桌台
 * @动作    批量写
 * @风险    高
 * @来源    store-data-extractor/PiLiangXiuGaiZuoTaiZiDingYiBiaoQian.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const axios = require('axios');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// Configuration
const INPUT_FILE = path.join(__dirname, '../input/需要修改的桌台.xlsx');
const BATCH_SIZE = 10;
const DELAY_BETWEEN_BATCHES = 1000; // 2 seconds

// Headers for the API calls
const HEADERS = {
    'accept': '*/*',
    'accept-language': 'zh-CN,zh;q=0.9',
    'authorization': 'bearer __VINCI_TOKEN__',
    'origin': 'https://connect.feibing.tech',
    'referer': 'https://connect.feibing.tech/',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/143.0.0.0 Safari/537.36',
    'cookie': 'x-token=__VINCI_TOKEN__; acw_tc=0a0f701417690526997471056e67c3d31605fd3fe8a78cc96fe2bfec88a70d' // Note: Used cookie from GET request as a baseline, but PUT request had a slightly different one in example. Using the most complete one or the one provided for PUT.
    // The user provided different cookies for GET and PUT. I'll use the one from PUT for the PUT request and GET for GET if needed, or just try to unify. 
    // Actually, usually the token is what matters. 
    // Let's use the PUT cookie for PUT request specifically if possible, or update HEADERS dynamically.
    // I'll stick to a common set and override if needed.
};

// PUT Request specific headers if needed (Content-Type is essential)
const PUT_HEADERS = {
    ...HEADERS,
    'Content-Type': 'application/json',
    'Cookie': 'x-token=__VINCI_TOKEN__; acw_tc=0a0f705417690508619526659e47fddae1e8e8ecb426729e04489896b392c7'
};

async function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

async function processPlaceId(id) {
    const getUrl = `https://vinci-api.feibing.tech/sc/v1/sellers/wwa9c5a585540b115b/places/${id}`;
    
    try {
        // 1. Get Place Details
        console.log(`正在获取详情: ${id}`);
        const getResponse = await axios.get(getUrl, { headers: HEADERS });
        
        if (getResponse.data.code !== 0) {
            console.error(`获取详情失败 ${id}: ${getResponse.data.message}`);
            return;
        }

        const metadata = getResponse.data.data.metadata;
        if (!metadata || !metadata.contactDerived) {
            console.error(`未找到 contactDerived 字段: ${id}`);
            return;
        }

        let contactDerivedStr = metadata.contactDerived;
        let contactDerivedObj;
        try {
            contactDerivedObj = JSON.parse(contactDerivedStr);
        } catch (e) {
            console.error(`解析 contactDerived 失败 ${id}: ${e.message}`);
            return;
        }

        // 2. Modify wxCpTags
        if (contactDerivedObj.wxCpTags) {
            contactDerivedObj.wxCpTags = [];
            console.log(`已清空 wxCpTags`);
        } else {
            console.log(`contactDerived 中无 wxCpTags, 跳过修改? (或者强制添加为空数组?) -> 按照需求是修改为空数组`);
             contactDerivedObj.wxCpTags = [];
        }

        const newContactDerivedStr = JSON.stringify(contactDerivedObj);

        // 3. Update Place Metadata
        const putUrl = `https://vinci-api.feibing.tech/sc/v1/sellers/wwa9c5a585540b115b/places/${id}/metadata/contactDerived`;
        console.log(`正在更新数据: ${id}`);
        
        // 修改：不再进行二次序列化，直接发送对象，让 axios 自动处理为 JSON 格式
        // 之前的写法 JSON.stringify(JSON.stringify(obj)) 会导致发送的内容变成 "{\"a\":1}" (带反斜杠的字符串)
        // 现在直接发送 obj，axios 会发送 {"a":1} (JSON对象)
        const payload = contactDerivedObj;
        
        const response = await axios.put(putUrl, payload, { 
            headers: PUT_HEADERS 
        });

        if (response.data.code === 0 || response.status === 200) {
            console.log(`更新成功: ${id}`);
        } else {
            console.error(`更新失败 ${id}: ${JSON.stringify(response.data)}`);
        }

    } catch (error) {
        console.error(`处理异常 ${id}:`, error.message);
        if (error.response) {
            console.error(`Status: ${error.response.status}`);
            console.error('Data:', error.response.data);
        }
    }
}

async function main() {
    if (!fs.existsSync(INPUT_FILE)) {
        console.error(`文件不存在: ${INPUT_FILE}`);
        return;
    }

    const workbook = XLSX.readFile(INPUT_FILE);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    // Assuming the column name is '桌台ID' or we take the first column if not found
    let ids = [];
    if (jsonData.length > 0) {
        if (jsonData[0]['桌台ID']) {
            ids = jsonData.map(row => row['桌台ID']).filter(id => id);
        } else {
            // Fallback: try to find a column that looks like an ID
             const keys = Object.keys(jsonData[0]);
             console.log(`未找到'桌台ID'列，尝试使用第一列: ${keys[0]}`);
             ids = jsonData.map(row => row[keys[0]]).filter(id => id);
        }
    }

    console.log(`共读取到 ${ids.length} 个桌台ID`);

    // Chunking and Batch Processing
    for (let i = 0; i < ids.length; i += BATCH_SIZE) {
        const batch = ids.slice(i, i + BATCH_SIZE);
        console.log(`\n--- 开始处理批次 ${Math.floor(i / BATCH_SIZE) + 1} (${batch.length}个) ---`);
        
        const promises = batch.map(id => processPlaceId(id));
        await Promise.all(promises);

        if (i + BATCH_SIZE < ids.length) {
            console.log(`批次处理完毕，等待 ${DELAY_BETWEEN_BATCHES}ms ...`);
            await sleep(DELAY_BETWEEN_BATCHES);
        }
        console.log(`批次处理完毕，继续下一批...`);
    }

    console.log('\n全部处理完成');
}

main().catch(console.error);

