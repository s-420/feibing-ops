/**
 * @对象    渠道码
 * @动作    批量写
 * @风险    高
 * @来源    store-data-extractor/PiLiangChuangJianQuDaoMa.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const axios = require('axios');
const fs = require('fs');


//   使用方法
/**
 * 在飞冰后台 沪上企业搜索空白码承接门店  切换到这个门店渠道桌台管理
 * 每创建一批空白码 都要创建一个分类  名称为年份+月份
 * 然后切换到新的分类下创建一个桌台 复制它的curl
 * 复制脚本 给AI让它替换curl请求  
 * 然后手动添加 "hideShopName": true,  切换type为CHANNEL
 * 
 * BASE_NAME创建的码的前缀 根据月份+日期定义一批码的名称  
 * START_INDEX创建码最后一位的起始序号  如果等于1   就是 3-24-1   然后根据  EXECUTE_TIMES去自增START_INDEX
 * 
 * CONCURRENCY并发数量   该脚本为并发任务队列形式   一次并发10条 有成功的就会自动触发下一条  不是固定10条全部请求完才能请求下一跳
 * MAX_RETRIES 并发内的每条请求可以重试的次数
 * RETRY_DELAY_MS 重试间隔（毫秒）
 * REQUEST_DELAY_MS 每个请求完成后额外延迟（毫秒），防止瞬时压力
 * FAILED_LOG_FILE  重试次数完了但是接口还没成功  错误日志抛出的文件名称
 */

// ========== 可配置参数 ==========
const EXECUTE_TIMES = 17758;           // 总请求数
const BASE_NAME = "3-24";              // 名称前缀
const START_INDEX = 2243;              // 起始序号
const CONCURRENCY = 10;                // 并发数（根据服务器压力调整）
const MAX_RETRIES = 3;                 // 单个请求最大重试次数（包含第一次尝试）
const RETRY_DELAY_MS = 1000;           // 重试间隔（毫秒）
const REQUEST_DELAY_MS = 200;          // 每个请求完成后额外延迟（毫秒），防止瞬时压力
const FAILED_LOG_FILE = 'failed.log';  // 失败任务日志文件
// ================================

const URL = 'https://vinci-api.feibing.tech/sc/v1/sellers/wwa9c5a585540b115b/places';

// 公共请求头
const HEADERS = {
  'accept': 'application/json',
  'accept-language': 'zh-CN,zh;q=0.9',
  'authorization': 'bearer __VINCI_TOKEN__',
  'content-type': 'application/json;charset=UTF-8',
  'origin': 'http://192.168.110.75:1024',
  'referer': 'http://192.168.110.75:1024/',
  'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36',
  'sec-ch-ua': '"Chromium";v="146", "Not-A.Brand";v="24", "Google Chrome";v="146"',
  'sec-ch-ua-mobile': '?0',
  'sec-ch-ua-platform': '"macOS"',
  'sec-fetch-dest': 'empty',
  'sec-fetch-mode': 'cors',
  'sec-fetch-site': 'cross-site',
  'sec-fetch-storage-access': 'active',
  'priority': 'u=1, i'
};

// 原始请求体
const baseData = {
  "hideShopName": true,
  "shopId": "69746ecc25056405fa738033",
  "$catalogs": "69c280d94841606c242ca269",
  "content": "Hello，可爱的新朋友\n🎁首次入群见面礼请查收\n\n当日『立减3元起』 \n次日❶最高减10元\n    ❷新品5️⃣折喝\n──────────\n🎁 更优福利天天享\n🔸多平台更优福利\n🔸赢海量饮品『免单』 \n\n进群👇🏻解锁更多福利",
  "table": {
    "type": "platformWeapp",
    "linkTrackConfig": {
      "urlTitle": "点我立即点单",
      "urlDesc": "",
      "urlPic": "https://boshsay.feibing.tech/20250225/affcdbeec29e471db24fbdc466da6261.png",
      "advance": true
    },
    "linkExpire": 0,
    "noticeType": 1,
    "appId": "wxd92a2d29f8022f40",
    "originalId": "gh_6613a6cac7ea",
    "path": "pages/index/index"
  },
  "COVER": [
    {
      "mid": "20260228287573080621056",
      "type": "IMAGE",
      "url": "https://boshsay.feibing.tech/20250225/affcdbeec29e471db24fbdc466da6261.png",
      "layout": "COVER",
      "title": "",
      "text": "",
      "order": 0
    }
  ],
  "attachments": [
    {
      "type": "link",
      "noticeType": 0,
      "link": {
        "title": "点击入群",
        "url": "https://work.weixin.qq.com/gm/1863c2a2ca4060151575c5c240e337fa",
        "desc": "解锁更多社群专属权益！",
        "linkExpire": 0,
        "picUrl": "https://fbhsay.su.bcebos.com/20240403/ef411ab289bb4bb68e66b1cfd7b36169.jpg"
      }
    }
  ],
  "contactDTO": {
    "pushApp": 1,
    "aggregateH5Valid": 1,
    "isOpenPreH5": 1,
    "preFilter": {
      "openFilter": false,
      "filterUserIds": []
    },
    "skipRelate": 0,
    "preH5Valid": 1,
    "isCreateClue": 0,
    "owners": ["15021108706"],
    "zfbRedirect": "https://ds.alipay.com/?scheme=alipays%3A%2F%2Fplatformapi%2Fstartapp%3FappId%3D2021002149651512%26page%3Dpages%2Ftakefood%2Findex%26query%3Dmulti_id%3D57621",
    "name": "空白码承接门店_测试创建",
    "tagSwitch": false,
    "wxCpTags": [],
    "qrcodeAvatar": "https://boshsay.feibing.tech/20250225/affcdbeec29e471db24fbdc466da6261.png",
    "welMsg": {
      "attachments": [
        {
          "type": "platformWeapp",
          "linkTrackConfig": {
            "urlTitle": "点我立即点单",
            "urlDesc": "",
            "urlPic": "https://boshsay.feibing.tech/20250225/affcdbeec29e471db24fbdc466da6261.png",
            "advance": true
          },
          "linkExpire": 0,
          "noticeType": 1,
          "appId": "wxd92a2d29f8022f40",
          "originalId": "gh_6613a6cac7ea",
          "path": "pages/index/index"
        },
        {
          "type": "link",
          "link": "https://work.weixin.qq.com/gm/1863c2a2ca4060151575c5c240e337fa",
          "linkExpire": 0,
          "noticeType": 0,
          "linkTrackConfig": {
            "urlPic": "https://fbhsay.su.bcebos.com/20240403/ef411ab289bb4bb68e66b1cfd7b36169.jpg",
            "urlTitle": "点击入群",
            "urlDesc": "解锁更多社群专属权益！"
          }
        }
      ],
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
      "background": "https://fb-dev.oss-cn-shanghai.aliyuncs.com/699eda72e4b092e2477d29cc.jpg"
    }
  },
  "tagCheckBox": [],
  "posterId": "default",
  "tableStickersId": "6951e4bfde351d2759f776cf",
  "name": "测试创建",
  "type": "CHANNEL",
  "media": [
    {
      "mid": "20260228287573080621056",
      "type": "IMAGE",
      "url": "https://boshsay.feibing.tech/20250225/affcdbeec29e471db24fbdc466da6261.png",
      "layout": "COVER",
      "title": "",
      "text": "",
      "order": 0
    }
  ],
  "metadata": {
    "tableStickers": "{\"id\":\"6951e4bfde351d2759f776cf\",\"title\":\"2025/12/29 咖啡门店号A4\",\"type\":\"TABLE_STICKERS\",\"posterVo\":{\"imgScale\":7.753125,\"scale\":1.0447400241837967,\"title\":\"2025/12/29 咖啡门店号A4\",\"qrCodeH\":810,\"qrCodeW\":810,\"qrCodeX\":840,\"qrCodeY\":955,\"tableSize\":0,\"tableY\":0,\"tableXelementLayoutType\":\"ABSOLUTELY\",\"fontColor\":\"BLACK\",\"backgroundUrl\":\"https://boshsay.feibing.tech/20251229/7b03351807a845d0bf7c401592ece75d.jpg\",\"tableX\":0},\"items\":[{\"sn\":\"1002\",\"relativeSn\":\"\",\"xelementLayoutType\":\"ABSOLUTELY\",\"yelementLayoutType\":\"ABSOLUTELY\",\"elementImageType\":\"RECTANGLE\",\"elementMediaType\":\"TEXT\",\"color\":{\"r\":0,\"g\":0,\"b\":0},\"font\":{\"elementFontStyle\":1,\"fontSize\":0},\"x\":0,\"y\":0}]}",
    "poster": "{\"id\":\"default\",\"title\":\"默认样式\",\"type\":\"POSTER\",\"posterVo\":{\"id\":\"default\",\"title\":\"默认样式\",\"type\":\"POSTER\",\"backgroundUrl\":\"https://fb-dev.oss-cn-shanghai.aliyuncs.com/20230711/79dc2da81da64a07b4d91573a8661198.png\",\"qrCodeX\":225,\"qrCodeY\":618,\"qrCodeW\":300,\"qrCodeH\":300,\"tableX\":0,\"tableY\":0,\"tableSize\":0,\"imgScale\":2.34375,\"fontColor\":\"WHITE\",\"tableXelementLayoutType\":\"CENTER\",\"background\":\"https://fb-dev.oss-cn-shanghai.aliyuncs.com/699eda72e4b092e2477d29cc.jpg\"},\"items\":[{\"sn\":\"1002\",\"relativeSn\":\"\",\"xelementLayoutType\":\"CENTER\",\"yelementLayoutType\":\"ABSOLUTELY\",\"elementImageType\":\"RECTANGLE\",\"elementMediaType\":\"TEXT\",\"elementContent\":\"测试创建\",\"color\":{\"r\":255,\"g\":255,\"b\":255},\"font\":{\"elementFontStyle\":1,\"fontSize\":0},\"x\":0,\"y\":0}]}",
    "customerService": "{\"id\":\"6985914a250564517db5ef85\",\"name\":\"15021108706\",\"nickName\":\"15021108706\",\"avatar\":\"https://rescdn.qqmail.com/node/wwmng/wwmng/style/images/independent/DefaultAvatar$73ba92b5.png\",\"seller\":{\"id\":\"wwa9c5a585540b115b\",\"name\":\"沪上阿姨\"},\"phone\":\"\",\"email\":\"\",\"status\":\"ON\",\"statusName\":\"在职\",\"position\":\"WAITER\",\"positionName\":\"服务员\",\"department\":{\"id\":\"685c27c1c3242c1fc836b9e6\",\"name\":\"3213133919\"},\"metadata\":{\"wxUserid\":\"15021108706\",\"alias\":\"总部福利官-木薯\",\"wxCpStatus\":\"1\",\"position\":\"WAITER\",\"isLeader\":\"0\"},\"modifiedBy\":\"system\",\"createdTime\":1770361162843,\"modifiedTime\":1774287473439,\"tags\":[]}"
  },
  "catalog": {
    "name": "2026-3",
    "id": "69c280d94841606c242ca269"
  },
  "poster": {
    "id": "default",
    "title": "默认样式",
    "type": "POSTER",
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
      "background": "https://fb-dev.oss-cn-shanghai.aliyuncs.com/699eda72e4b092e2477d29cc.jpg"
    },
    "items": [
      {
        "sn": "1002",
        "relativeSn": "",
        "xelementLayoutType": "CENTER",
        "yelementLayoutType": "ABSOLUTELY",
        "elementImageType": "RECTANGLE",
        "elementMediaType": "TEXT",
        "elementContent": "测试创建",
        "color": { "r": 255, "g": 255, "b": 255 },
        "font": { "elementFontStyle": 1, "fontSize": 0 },
        "x": 0,
        "y": 0
      }
    ]
  },
  "tableStickers": {
    "id": "6951e4bfde351d2759f776cf",
    "title": "2025/12/29 咖啡门店号A4",
    "type": "TABLE_STICKERS",
    "posterVo": {
      "imgScale": 7.753125,
      "scale": 1.0447400241837967,
      "title": "2025/12/29 咖啡门店号A4",
      "qrCodeH": 810,
      "qrCodeW": 810,
      "qrCodeX": 840,
      "qrCodeY": 955,
      "tableSize": 0,
      "tableY": 0,
      "tableXelementLayoutType": "ABSOLUTELY",
      "fontColor": "BLACK",
      "backgroundUrl": "https://boshsay.feibing.tech/20251229/7b03351807a845d0bf7c401592ece75d.jpg",
      "tableX": 0
    },
    "items": [
      {
        "sn": "1002",
        "relativeSn": "",
        "xelementLayoutType": "ABSOLUTELY",
        "yelementLayoutType": "ABSOLUTELY",
        "elementImageType": "RECTANGLE",
        "elementMediaType": "TEXT",
        "color": { "r": 0, "g": 0, "b": 0 },
        "font": { "elementFontStyle": 1, "fontSize": 0 },
        "x": 0,
        "y": 0
      }
    ]
  }
};

// 辅助函数：深度替换对象中所有匹配 oldStr 的字符串为 newStr
function replaceAllStrings(obj, oldStr, newStr) {
  if (typeof obj === 'string') {
    return obj.replace(new RegExp(oldStr, 'g'), newStr);
  }
  if (Array.isArray(obj)) {
    return obj.map(item => replaceAllStrings(item, oldStr, newStr));
  }
  if (obj && typeof obj === 'object') {
    const newObj = {};
    for (const key in obj) {
      newObj[key] = replaceAllStrings(obj[key], oldStr, newStr);
    }
    return newObj;
  }
  return obj;
}

// 生成请求体
function buildRequestBody(name) {
  return replaceAllStrings(baseData, '测试创建', name);
}

// 带重试的单次请求
async function sendRequestWithRetry(name, index) {
  const data = buildRequestBody(name);
  let lastError = null;
  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      console.log(`[${index}] 发送请求 (${attempt}/${MAX_RETRIES})，名称: ${name}`);
      const response = await axios.post(URL, data, { headers: HEADERS, timeout: 30000 });
      console.log(`[${index}] 请求成功，名称: ${name}`);
      return response.data;
    } catch (error) {
      lastError = error;
      console.error(`[${index}] 请求失败 (${attempt}/${MAX_RETRIES})，错误: ${error.message}`);
      if (attempt < MAX_RETRIES) {
        await new Promise(resolve => setTimeout(resolve, RETRY_DELAY_MS));
      }
    }
  }
  // 重试耗尽，记录失败
  const failMsg = `[${index}] ${name} 失败，最终错误: ${lastError.message}\n`;
  console.error(failMsg);
  fs.appendFileSync(FAILED_LOG_FILE, failMsg);
  throw lastError; // 抛出以便队列感知失败，但不停止整体
}

// 并发队列控制
class ConcurrencyQueue {
  constructor(concurrency, delayMs = 0) {
    this.concurrency = concurrency;
    this.delayMs = delayMs;
    this.running = 0;
    this.queue = [];
  }

  addTask(task) {
    return new Promise((resolve, reject) => {
      this.queue.push({ task, resolve, reject });
      this._run();
    });
  }

  _run() {
    while (this.running < this.concurrency && this.queue.length > 0) {
      const { task, resolve, reject } = this.queue.shift();
      this.running++;
      task()
        .then(result => {
          this.running--;
          resolve(result);
          if (this.delayMs > 0) {
            setTimeout(() => this._run(), this.delayMs);
          } else {
            this._run();
          }
        })
        .catch(err => {
          this.running--;
          reject(err);
          this._run();
        });
    }
  }

  async runAll(tasks) {
    const promises = tasks.map(task => this.addTask(task));
    return Promise.allSettled(promises);
  }
}

// 主函数
async function main() {
  // 可选：清空旧的失败日志
  // fs.writeFileSync(FAILED_LOG_FILE, '');

  console.log(`总任务数: ${EXECUTE_TIMES}`);
  console.log(`起始序号: ${START_INDEX}`);
  console.log(`并发数: ${CONCURRENCY}`);
  console.log(`单个请求最大重试次数: ${MAX_RETRIES}`);
  console.log(`请求后延迟: ${REQUEST_DELAY_MS}ms`);
  console.log(`失败日志文件: ${FAILED_LOG_FILE}`);
  console.log('开始执行...\n');

  // 生成所有任务
  const tasks = [];
  for (let i = 0; i < EXECUTE_TIMES; i++) {
    const currentIndex = START_INDEX + i;
    const fullName = `${BASE_NAME}-${currentIndex}`;
    tasks.push(() => sendRequestWithRetry(fullName, currentIndex));
  }

  const queue = new ConcurrencyQueue(CONCURRENCY, REQUEST_DELAY_MS);
  const results = await queue.runAll(tasks);

  // 统计结果
  let successCount = 0;
  let failCount = 0;
  results.forEach(result => {
    if (result.status === 'fulfilled') {
      successCount++;
    } else {
      failCount++;
    }
  });

  console.log(`\n执行完成！成功: ${successCount}, 失败: ${failCount}`);
  if (failCount > 0) {
    console.log(`失败详情请查看: ${FAILED_LOG_FILE}`);
  }
}

main().catch(console.error);