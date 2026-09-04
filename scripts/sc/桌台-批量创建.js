/**
 * @对象    桌台
 * @动作    建实体
 * @风险    中
 * @来源    store-data-extractor/XunHuanChuangJianZuoTai.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const axios = require('axios');
const path = require('path');

// 配置项
const config = {
  // API 配置
  apiUrl: 'https://vinci-api.feibing.tech/sc/v1/sellers/wwa9c5a585540b115b/places',
  headers: {
    'accept': 'application/json',
    'accept-language': 'zh-CN,zh;q=0.9',
    'authorization': 'bearer __VINCI_TOKEN__',
    'content-type': 'application/json;charset=UTF-8',
    'origin': 'https://connect.feibing.tech',
    'priority': 'u=1, i',
    'referer': 'https://connect.feibing.tech/',
    'sec-ch-ua': '"Not(A:Brand";v="8", "Chromium";v="144", "Google Chrome";v="144"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"macOS"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-site',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36'
  },
  cookies: 'x-token=__VINCI_TOKEN__; acw_tc=0a0f705317697469281231952e60c0ee420972e9f593ff26635b90e7a87cf2',
  
  // 循环控制配置
  maxCalls: 119,              // 最大调用次数
  startNumber: 4,            // name 字段起始数字（从 "1-30-3" 中的最后一个数字开始）
  namePrefix: '1-30-',       // name 字段前缀（"1-30-"）
  delayBetweenCalls: 1000,   // 每次调用之间的延迟（毫秒）
  
  // 请求超时配置
  timeout: 30000
};

// 基础请求体模板（从 curl 命令中提取）
const baseRequestBody = {
  "shopId": "69746ecc25056405fa738033",
  "$catalogs": "697af7af0a187005ecc7500c",
  "content": "Hello，可爱的新朋友\n🎁首次入群见面礼请查收\n\n当日『立减3元起』 \n次日❶最高减10元\n    ❷新品5️⃣折喝\n──────────\n🎁 更优福利天天享\n🔸多平台更优福利\n🔸赢海量饮品『免单』 \n\n进群👇🏻解锁更多福利",
  "table": {
    "type": "platformWeapp",
    "linkTrackConfig": {
      "urlTitle": "点我立即点单",
      "urlDesc": "",
      "urlPic": "https://boshsay.feibing.tech/20240322/38a722af4e9d4e18bda10a8476ebdaac.jpg",
      "advance": false,
      "activeNotice": false,
      "behaviorNotice": false,
      "saveAsRadarLink": false
    },
    "linkExpire": 0,
    "noticeType": 1,
    "link": "",
    "originalId": "gh_6613a6cac7ea",
    "path": "pages/takefood/index?multi_id=57621",
    "appId": "wxd92a2d29f8022f40"
  },
  "COVER": [{
    "mid": "20260130185383435452417",
    "type": "IMAGE",
    "url": "https://boshsay.feibing.tech/20240322/38a722af4e9d4e18bda10a8476ebdaac.jpg",
    "layout": "COVER",
    "title": "",
    "text": "",
    "order": 0
  }],
  "attachments": [{
    "type": "link",
    "link": {
      "url": "https://work.weixin.qq.com/gm/1863c2a2ca4060151575c5c240e337fa",
      "title": "点击入群",
      "desc": "解锁更多社群专属权益！",
      "picUrl": "https://boshsay.feibing.tech/20240403/ef411ab289bb4bb68e66b1cfd7b36169.jpg",
      "linkExpire": 0
    },
    "noticeType": 0
  }],
  "contactDTO": {
    "pushApp": 1,
    "aggregateH5Valid": 1,
    "isOpenPreH5": 1,
    "preFilter": {
      "openFilter": false,
      "filterUserIds": [],
      "$filterUserIds": false
    },
    "skipRelate": 0,
    "preH5Valid": 1,
    "isCreateClue": 0,
    "owners": ["13122435505"],
    "wxCpTags": [],
    "zfbRedirect": "https://ds.alipay.com/?scheme=alipays%3A%2F%2Fplatformapi%2Fstartapp%3FappId%3D2021002149651512%26page%3Dpages%2Ftakefood%2Findex%26query%3Dmulti_id%3Dnull",
    "name": "空白码承接门店_1-30-3",
    "tagSwitch": false,
    "qrcodeAvatar": "https://boshsay.feibing.tech/20240322/38a722af4e9d4e18bda10a8476ebdaac.jpg",
    "welMsg": {
      "attachments": [{
        "type": "platformWeapp",
        "linkTrackConfig": {
          "urlTitle": "点我立即点单",
          "urlDesc": "",
          "urlPic": "https://boshsay.feibing.tech/20240322/38a722af4e9d4e18bda10a8476ebdaac.jpg",
          "advance": false,
          "activeNotice": false,
          "behaviorNotice": false,
          "saveAsRadarLink": false
        },
        "linkExpire": 0,
        "noticeType": 1,
        "link": "",
        "originalId": "gh_6613a6cac7ea",
        "path": "pages/takefood/index?multi_id=57621",
        "appId": "wxd92a2d29f8022f40"
      }, {
        "type": "link",
        "link": "https://work.weixin.qq.com/gm/1863c2a2ca4060151575c5c240e337fa",
        "linkExpire": 0,
        "noticeType": 0,
        "linkTrackConfig": {
          "urlPic": "https://boshsay.feibing.tech/20240403/ef411ab289bb4bb68e66b1cfd7b36169.jpg",
          "urlTitle": "点击入群",
          "urlDesc": "解锁更多社群专属权益！"
        }
      }],
      "content": "Hello，可爱的新朋友\n🎁首次入群见面礼请查收\n\n当日『立减3元起』 \n次日❶最高减10元\n    ❷新品5️⃣折喝\n──────────\n🎁 更优福利天天享\n🔸多平台更优福利\n🔸赢海量饮品『免单』 \n\n进群👇🏻解锁更多福利"
    },
    "posterVo": {
      "id": "default",
      "title": "默认样式",
      "type": "POSTER",
      "backgroundUrl": "https://fb-dev.oss-cn-shanghai.aliyuncs.com/20230711/79dc2da81da64a07b4d91573a8661198.png",
      "qrCodeX": 225,
      "qrCodeY": 618,
      "qrCodeW": 300,
      "qrCodeH": 300,
      "tableX": 0,
      "tableY": 0,
      "tableSize": 0,
      "imgScale": 2.34375,
      "fontColor": "WHITE",
      "tableXelementLayoutType": "CENTER",
      "background": "https://fb-dev.oss-cn-shanghai.aliyuncs.com/697c3223e4b092e20612394f.jpg"
    }
  },
  "tagCheckBox": [],
  "posterId": "default",
  "tableStickersId": "6951e4bfde351d2759f776cf",
  "seller": {
    "id": "wwa9c5a585540b115b",
    "name": "沪上阿姨"
  },
  "shop": {
    "id": "69746ecc25056405fa738033",
    "name": "空白码承接门店",
    "branchName": ""
  },
  "catalog": {
    "name": "2026-1",
    "id": "697af7af0a187005ecc7500c"
  },
  "name": "1-30-3",
  "description": "",
  "type": "CHANNEL",
  "userLimit": -1,
  "status": "IDLE",
  "statusName": "闲置",
  "media": [{
    "mid": "20260130185383435452417",
    "type": "IMAGE",
    "url": "https://boshsay.feibing.tech/20240322/38a722af4e9d4e18bda10a8476ebdaac.jpg",
    "layout": "COVER",
    "title": "",
    "text": "",
    "order": 0
  }]
};

// 生成 name 字段
function generateName(number) {
  return `${config.namePrefix}${number}`;
}

// 更新请求体中的 name 字段（包括所有相关字段）
function updateRequestBodyName(requestBody, name) {
  // 深拷贝请求体
  const updated = JSON.parse(JSON.stringify(requestBody));
  
  // 更新主 name 字段
  updated.name = name;
  
  // 更新 contactDTO.name 字段（如果存在）
  if (updated.contactDTO) {
    const shopName = updated.shop?.name || '空白码承接门店';
    updated.contactDTO.name = `${shopName}_${name}`;
  }
  
  return updated;
}

// 调用 API
async function callAPI(requestBody, callNumber) {
  try {
    const headers = {
      ...config.headers,
      'Cookie': config.cookies
    };

    console.log(`📤 第 ${callNumber} 次调用，name: ${requestBody.name}`);
    
    const response = await axios.post(config.apiUrl, requestBody, {
      headers: headers,
      timeout: config.timeout
    });

    // 检查响应
    if (response.status === 200 || response.status === 201) {
      const responseData = response.data;
      
      // 如果返回的是 JSON 对象，检查是否有 code 字段
      if (responseData && typeof responseData === 'object') {
        if (responseData.code === 0 || responseData.code === undefined) {
          console.log(`✅ 第 ${callNumber} 次调用成功，name: ${requestBody.name}`);
          return { success: true, data: responseData, callNumber };
        } else {
          console.log(`⚠️ 第 ${callNumber} 次调用返回异常，code: ${responseData.code}, message: ${responseData.message || '未知错误'}`);
          return { success: false, error: `API返回异常: code=${responseData.code}, message=${responseData.message}`, callNumber };
        }
      } else {
        console.log(`✅ 第 ${callNumber} 次调用成功，name: ${requestBody.name}`);
        return { success: true, data: responseData, callNumber };
      }
    } else {
      console.log(`⚠️ 第 ${callNumber} 次调用返回状态码: ${response.status}`);
      return { success: false, error: `HTTP状态码: ${response.status}`, callNumber };
    }
  } catch (error) {
    const errorMsg = error.response?.data?.message || error.message || '未知错误';
    console.log(`❌ 第 ${callNumber} 次调用失败，name: ${requestBody.name}, 错误: ${errorMsg}`);
    return { success: false, error: errorMsg, callNumber };
  }
}

// 延迟函数
function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// 主函数
async function main() {
  console.log('🚀 开始循环调用接口');
  console.log(`📋 配置信息:`);
  console.log(`   - 最大调用次数: ${config.maxCalls}`);
  console.log(`   - name 前缀: ${config.namePrefix}`);
  console.log(`   - 起始数字: ${config.startNumber}`);
  console.log(`   - 调用间隔: ${config.delayBetweenCalls}ms`);
  console.log('------------------------------\n');

  const results = {
    total: 0,
    success: 0,
    failed: 0,
    details: []
  };

  for (let i = 0; i < config.maxCalls; i++) {
    const currentNumber = config.startNumber + i;
    const name = generateName(currentNumber);
    const requestBody = updateRequestBodyName(baseRequestBody, name);
    
    results.total++;
    
    const result = await callAPI(requestBody, i + 1);
    
    if (result.success) {
      results.success++;
    } else {
      results.failed++;
    }
    
    results.details.push({
      callNumber: i + 1,
      name: name,
      success: result.success,
      error: result.error || null
    });

    // 如果不是最后一次调用，等待延迟
    if (i < config.maxCalls - 1) {
      await delay(config.delayBetweenCalls);
    }
  }

  // 输出统计结果
  console.log('\n------------------------------');
  console.log('📊 执行结果统计:');
  console.log(`   总调用次数: ${results.total}`);
  console.log(`   成功次数: ${results.success} (${((results.success / results.total) * 100).toFixed(2)}%)`);
  console.log(`   失败次数: ${results.failed} (${((results.failed / results.total) * 100).toFixed(2)}%)`);
  
  if (results.failed > 0) {
    console.log('\n❌ 失败详情:');
    results.details
      .filter(r => !r.success)
      .forEach(r => {
        console.log(`   第 ${r.callNumber} 次调用 (name: ${r.name}): ${r.error}`);
      });
  }
  
  console.log('\n🎉 脚本执行完成');
}

// 启动脚本
main().catch(error => {
  console.error('💥 脚本执行出错:', error.message);
  process.exit(1);
});

