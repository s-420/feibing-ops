/**
 * @对象    优惠券
 * @动作    批量写
 * @风险    高
 * @来源    store-data-extractor/XunHuanDiaoYongFaQuanLianLu.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

// 调用命令
// npm run XunHuanDiaoYongFaQuanLianLu -- --count=10 --delay=100
// send 配置请直接修改下方 SEND_ITEMS 常量

const axios = require("axios");
const fs = require("fs");
const path = require("path");

const AUTHORIZATION =
  process.env.FEIBING_AUTHORIZATION ||
  "bearer __VINCI_TOKEN__";

const OUTPUT_DIR = path.resolve(__dirname, "../output");
const DEFAULT_TIMEOUT = 15000;
const SEND_ITEMS = [
  {
    ticket_id: 1285759,
    total: 1,
  },
  {
    ticket_id: 1285760,
    total: 2,
  },
  {
    ticket_id: 1285761,
    total: 2,
  },
  {
    ticket_id: 1285762,
    total: 2,
  },
];

const BASE_HEADERS = {
  Connection: "keep-alive",
  Pragma: "no-cache",
  "Cache-Control": "no-cache",
  authorization: AUTHORIZATION,
  "X-Scene": "1194",
  "X-Group-Encrypted": "",
  "User-Agent":
    "Mozilla/5.0 (iPhone; CPU iPhone OS 11_0 like Mac OS X) AppleWebKit/604.1.38 (KHTML, like Gecko) Version/11.0 Mobile/15A372 Safari/604.1 wechatdevtools/1.06.2504060 MicroMessenger/8.0.5 Language/zh_CN webview/ sessionid/201",
  "X-Group-Iv": "",
  "content-type": "application/json",
  Accept: "*/*",
  "Sec-Fetch-Site": "cross-site",
  "Sec-Fetch-Mode": "cors",
  "Sec-Fetch-Dest": "empty",
  Referer: "https://servicewechat.com/wx3c353dd9715ac456/devtools/page-frame.html",
  "Accept-Language": "zh-CN,zh;q=0.9",
};

const PARTICIPANTS_BODY = {
  activityId: "69c0f1568e922c063e89f3b6",
  sellerId: "wpV3cNDAAA7gXDSpALtCGXgFxSpYMBQQ",
  user: {
    id: "682adb11e3725c06d0f11a69",
    login: "oHw2bt3i1gii8cTNYcPqf2cnHuUs",
    realName: "",
    gender: "UNKNOWN",
    nickName: "i",
    avatar:
      "https://thirdwx.qlogo.cn/mmopen/vi_32/gjLmQ3YdADiaT3xDUGrN8cXLIWtJ5Z7Zn1daBombY8PQc5ORfVzCSib0sLFBvduw0KkaWQvXdP1TCcVQic1uH9PtFb7TWQff2kYqicibI1MbKQHA/132",
    activated: true,
    langKey: "cn",
    authorities: ["ROLE_USER"],
    wx: {
      nickName: "i",
      gender: "UNKNOWN",
      language: "",
      province: "",
      country: "",
      city: "",
      avatar:
        "https://thirdwx.qlogo.cn/mmopen/vi_32/gjLmQ3YdADiaT3xDUGrN8cXLIWtJ5Z7Zn1daBombY8PQc5ORfVzCSib0sLFBvduw0KkaWQvXdP1TCcVQic1uH9PtFb7TWQff2kYqicibI1MbKQHA/132",
      unionId: "oHw2bt3i1gii8cTNYcPqf2cnHuUs",
      openId: "oIcqA5Zmlm9C-l-zXlHEClST7cXU",
    },
    registerType: "wx_mini_app_register",
    registerIp: "61.171.250.34",
    createdBy: "web_app",
    modifiedBy: "web_app",
    createdTime: 1747639057064,
    modifiedTime: 1772778833342,
    isAnonymous: false,
    metadata: {
      seller_list:
        '[{"id":"wwa0eca081a49d28e9","logo":"https://fb-dev.oss-cn-shanghai.aliyuncs.com/20220523/337928adaa0f4f1ca1ddb120b5370f2e.png","status":"OPEN","name":"飞冰科技","statusName":"营业","type":"FACILITATOR","description":"飞冰科技"},{"id":"wpV3cNDAAAa5pLikCaI4KbhtYUxy9whw","logo":"https://p.qlogo.cn/bizmail/xu0ia48ibYyVts01j0JErZWEdhicL3tPTnAvqfW9RYibMgXvP1T9iaOcExA/0","status":"OPEN","name":"青岛啤酒电子商务有限公司","statusName":"营业","type":"FACILITATOR","description":"青岛啤酒电子商务有限公司"}]',
      wpV3cNDAAAom_b__MC0feu5vw67JNbrA_league_wxUser_appId: "wx554abb748dcab821",
      wpV3cNDAAAa5pLikCaI4KbhtYUxy9whw_league: "true",
      wpV3cNDAAAom_b__MC0feu5vw67JNbrA_league_grade: "0",
      wpV3cNDAAAa5pLikCaI4KbhtYUxy9whw_level: "0",
      wwa0eca081a49d28e9_league_grade: "1",
      wwa0eca081a49d28e9_level_coefficient: "10.0",
      wwa0eca081a49d28e9_league_wxUser_openId: "o3fjz4mHDV6WW2WK9WAjaZnseMow",
      wwa0eca081a49d28e9_league_language: "",
      wpV3cNDAAAom_b__MC0feu5vw67JNbrA_league_nickName: "i",
      wpV3cNDAAAa5pLikCaI4KbhtYUxy9whw_level_coefficient: "1.0",
      wpV3cNDAAAom_b__MC0feu5vw67JNbrA_league_language: "",
      league: "true",
      wwa0eca081a49d28e9_sharerIntroductionV3: "test1",
      wpV3cNDAAAa5pLikCaI4KbhtYUxy9whw_league_avatar:
        "https://thirdwx.qlogo.cn/mmopen/vi_32/gjLmQ3YdADiaT3xDUGrN8cXLIWtJ5Z7Zn1daBombY8PTVQSVc6VibwzHPvDgbYs4VMRFwYP0vBIeEdAgRlDJFUb5XATtRAw3arLeNIibx1Bqzw/132",
      wwa0eca081a49d28e9_league_avatar:
        "https://thirdwx.qlogo.cn/mmopen/vi_32/gjLmQ3YdADiaT3xDUGrN8cXLIWtJ5Z7Zn1daBombY8PTVQSVc6VibwzHPvDgbYs4VMRFwYP0vBIeEdAgRlDJFUb5XATtRAw3arLeNIibx1Bqzw/132",
      wpV3cNDAAAa5pLikCaI4KbhtYUxy9whw_league_sharerAppid: "wxa224244f1096344c",
      wwa0eca081a49d28e9_league: "true",
      wpV3cNDAAAom_b__MC0feu5vw67JNbrA_league_bindInfo:
        '{"errcode":0,"errmsg":"ok","bindStatus":1,"registerStatus":2,"registerBusinessType":"","registerQueryString":"","bindBusinessType":"","bindQueryString":""}',
      wpV3cNDAAAom_b__MC0feu5vw67JNbrA_league_gender: "UNKNOWN",
      wpV3cNDAAAom_b__MC0feu5vw67JNbrA_league_openBusinessView_extraData:
        '{"commissionType":1,"commissionRatio":100000,"headSupplierAppid":"wxdc4b703d4e014f60"}',
      wpV3cNDAAAa5pLikCaI4KbhtYUxy9whw_league_wxUser_appId: "wx3e6d74be56e59ae5",
      wwa0eca081a49d28e9_league_inviter: "",
      wpV3cNDAAAa5pLikCaI4KbhtYUxy9whw_league_inviter: "",
      wpV3cNDAAAom_b__MC0feu5vw67JNbrA_league_openfinderid: "",
      wpV3cNDAAAa5pLikCaI4KbhtYUxy9whw_league_openBusinessView_extraData:
        '{"commissionType":1,"commissionRatio":1000000,"headSupplierAppid":"wxb13ed60563429af0"}',
      wpV3cNDAAAom_b__MC0feu5vw67JNbrA_league_country: "",
      wpV3cNDAAAom_b__MC0feu5vw67JNbrA_league_openId: "oIcqA5Zmlm9C-l-zXlHEClST7cXU",
      wwa0eca081a49d28e9_league_sharerAppid: "wxa224244f1096344c",
      wpV3cNDAAAa5pLikCaI4KbhtYUxy9whw_league_bindInfo:
        '{"errcode":0,"errmsg":"ok","bindStatus":1,"registerStatus":2,"registerBusinessType":"","registerQueryString":"","bindBusinessType":"","bindQueryString":""}',
      wwa0eca081a49d28e9_league_bindInfo:
        '{"errcode":0,"errmsg":"ok","bindStatus":1,"registerStatus":2,"registerBusinessType":"","registerQueryString":"","bindBusinessType":"","bindQueryString":""}',
      wpV3cNDAAAom_b__MC0feu5vw67JNbrA_league_unionId: "oHw2bt3i1gii8cTNYcPqf2cnHuUs",
      wwa0eca081a49d28e9_league_openBusinessView_extraData:
        '{"commissionType":1,"commissionRatio":1000000,"headSupplierAppid":"wx07af89e138a0a424"}',
      wpV3cNDAAAa5pLikCaI4KbhtYUxy9whw_league_province: "",
      wwa0eca081a49d28e9_league_openId: "oIcqA5Zmlm9C-l-zXlHEClST7cXU",
      wwa0eca081a49d28e9_sharerPhoneV3: "",
      wwa0eca081a49d28e9_level: "0",
      wpV3cNDAAAom_b__MC0feu5vw67JNbrA_league_bindStatus: "1",
      wpV3cNDAAAom_b__MC0feu5vw67JNbrA_league_wxUser_openId: "obT6s7dzf7w6BeG7iNnqQ2CmdMWQ",
      wpV3cNDAAAom_b__MC0feu5vw67JNbrA_league: "true",
      wpV3cNDAAAom_b__MC0feu5vw67JNbrA_league_city: "",
      wwa0eca081a49d28e9_league_bindStatus: "1",
      wwa0eca081a49d28e9_league_unionId: "oHw2bt3i1gii8cTNYcPqf2cnHuUs",
      wpV3cNDAAAa5pLikCaI4KbhtYUxy9whw_league_wxUser_openId: "oTjEH7I3ZqvJI6yi4O-KxF44YmCc",
      wpV3cNDAAAa5pLikCaI4KbhtYUxy9whw_league_language: "",
      wpV3cNDAAAa5pLikCaI4KbhtYUxy9whw_league_nickName: "i",
      wpV3cNDAAAa5pLikCaI4KbhtYUxy9whw_league_bindStatus: "1",
      wpV3cNDAAAa5pLikCaI4KbhtYUxy9whw_league_unionId: "oHw2bt3i1gii8cTNYcPqf2cnHuUs",
      wpV3cNDAAAa5pLikCaI4KbhtYUxy9whw_league_grade: "1",
      wwa0eca081a49d28e9_league_nickName: "i",
      wpV3cNDAAAom_b__MC0feu5vw67JNbrA_league_sharerAppid: "wxa224244f1096344c",
      wpV3cNDAAAom_b__MC0feu5vw67JNbrA_league_finderid: "",
      wpV3cNDAAAa5pLikCaI4KbhtYUxy9whw_league_openId: "oIcqA5Zmlm9C-l-zXlHEClST7cXU",
      wwa0eca081a49d28e9_league_province: "",
      wpV3cNDAAAom_b__MC0feu5vw67JNbrA_league_province: "",
      wpV3cNDAAAom_b__MC0feu5vw67JNbrA_league_avatar:
        "https://fbbc.cdn.bcebos.com/20250718/35b63116de6d42a8893f751922194705.jpg",
      wwa0eca081a49d28e9_league_wxUser_appId: "wx64c652cf391b7183",
    },
    language: "",
    province: "",
    country: "",
    city: "",
    unionId: "oHw2bt3i1gii8cTNYcPqf2cnHuUs",
    openId: "oIcqA5Zmlm9C-l-zXlHEClST7cXU",
  },
  metadata: {},
};

const STATIC_CONFIG = {
  sellerId: "wpV3cNDAAA7gXDSpALtCGXgFxSpYMBQQ",
  unionId: "oHw2bt3i1gii8cTNYcPqf2cnHuUs",
  openId: "oIcqA5Zmlm9C-l-zXlHEClST7cXU",
  initialEid: "wmV3cNDAAACboJkfGSUBvUM0qE4o9HXw",
  couponOpenId: "oR9Y6wsmoh8sCTcquALLsqqhQfYg",
};

function parseArgs(argv) {
  const options = {
    count: Number(process.env.FEIBING_LOOP_COUNT || 1),
    delayMs: Number(process.env.FEIBING_DELAY_MS || 0),
    timeout: Number(process.env.FEIBING_TIMEOUT_MS || DEFAULT_TIMEOUT),
  };

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];

    if (arg === "--count" || arg === "-c") {
      options.count = Number(argv[index + 1]);
      index += 1;
      continue;
    }

    if (arg.startsWith("--count=")) {
      options.count = Number(arg.split("=")[1]);
      continue;
    }

    if (arg === "--delay" || arg === "-d") {
      options.delayMs = Number(argv[index + 1]);
      index += 1;
      continue;
    }

    if (arg.startsWith("--delay=")) {
      options.delayMs = Number(arg.split("=")[1]);
      continue;
    }

    if (arg === "--timeout") {
      options.timeout = Number(argv[index + 1]);
      index += 1;
      continue;
    }

    if (arg.startsWith("--timeout=")) {
      options.timeout = Number(arg.split("=")[1]);
      continue;
    }
  }

  if (!Number.isInteger(options.count) || options.count <= 0) {
    throw new Error("循环次数必须是大于 0 的整数，可通过 --count 指定。");
  }

  if (!Number.isInteger(options.delayMs) || options.delayMs < 0) {
    throw new Error("delayMs 必须是大于等于 0 的整数。");
  }

  if (!Number.isInteger(options.timeout) || options.timeout <= 0) {
    throw new Error("timeout 必须是大于 0 的整数。");
  }

  return options;
}

function ensureOutputDir() {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

function formatFileTime(date = new Date()) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hour = String(date.getHours()).padStart(2, "0");
  const minute = String(date.getMinutes()).padStart(2, "0");
  const second = String(date.getSeconds()).padStart(2, "0");
  return `${year}${month}${day}_${hour}${minute}${second}`;
}

function createRunFiles() {
  ensureOutputDir();
  const stamp = formatFileTime();
  return {
    logFile: path.join(OUTPUT_DIR, `faquan_loop_${stamp}.log`),
    errorFile: path.join(OUTPUT_DIR, `faquan_loop_${stamp}.error.log`),
    summaryFile: path.join(OUTPUT_DIR, `faquan_loop_${stamp}.json`),
  };
}

function appendLog(logFile, message) {
  const line = `[${new Date().toISOString()}] ${message}\n`;
  fs.appendFileSync(logFile, line, "utf8");
  console.log(message);
}

function writeSummary(summaryFile, summary) {
  fs.writeFileSync(summaryFile, `${JSON.stringify(summary, null, 2)}\n`, "utf8");
}

function appendErrorLog(errorFile, lines) {
  const message = Array.isArray(lines) ? lines.join("\n") : String(lines);
  const block = `[${new Date().toISOString()}] ${message}\n\n`;
  fs.appendFileSync(errorFile, block, "utf8");
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function generate24HexId() {
  const timestamp = Math.floor(Date.now() / 1000).toString(16);
  const timePart = timestamp.padStart(8, "0").slice(-8);

  let randomPart = "";
  const hexChars = "0123456789abcdef";
  for (let index = 0; index < 16; index += 1) {
    randomPart += hexChars[Math.floor(Math.random() * 16)];
  }

  return (timePart + randomPart).toLowerCase();
}

function createOrderRequestSn() {
  return generate24HexId();
}

function normalizeSendItems(sendItems) {
  if (!Array.isArray(sendItems) || sendItems.length === 0) {
    throw new Error("send 配置必须是非空数组，例如 [{\"ticket_id\":1285760,\"total\":2}]");
  }

  return sendItems.map((item, index) => {
    if (!item || typeof item !== "object") {
      throw new Error(`第 ${index + 1} 个 send 配置不是对象。`);
    }

    const ticketId = Number(item.ticket_id ?? item.ticketId);
    const total = Number(item.total);

    if (!Number.isInteger(ticketId) || ticketId <= 0) {
      throw new Error(`第 ${index + 1} 个 send 配置的 ticket_id 非法。`);
    }

    if (!Number.isInteger(total) || total <= 0) {
      throw new Error(`第 ${index + 1} 个 send 配置的 total 非法。`);
    }

    return {
      ticket_id: ticketId,
      total,
    };
  });
}

function safePreview(data) {
  if (data === undefined) {
    return "undefined";
  }

  if (data === null) {
    return "null";
  }

  if (typeof data === "string") {
    return data.slice(0, 500);
  }

  try {
    return JSON.stringify(data).slice(0, 500);
  } catch (error) {
    return `[无法序列化响应: ${error.message}]`;
  }
}

function evaluateBusinessResult(data) {
  if (data === null || data === undefined) {
    return { success: true, businessCode: null, message: "" };
  }

  if (typeof data !== "object") {
    return { success: true, businessCode: null, message: "" };
  }

  if (Object.prototype.hasOwnProperty.call(data, "success")) {
    return {
      success: Boolean(data.success),
      businessCode: data.code ?? null,
      message: data.message || data.msg || data.errmsg || "",
    };
  }

  if (Object.prototype.hasOwnProperty.call(data, "errcode")) {
    return {
      success: Number(data.errcode) === 0,
      businessCode: data.errcode,
      message: data.errmsg || data.message || "",
    };
  }

  if (Object.prototype.hasOwnProperty.call(data, "code")) {
    const code = data.code;
    const success = code === 0 || code === "0" || code === 200 || code === "200";
    return {
      success,
      businessCode: code,
      message: data.message || data.msg || data.errmsg || "",
    };
  }

  if (typeof data.status === "string") {
    const normalizedStatus = data.status.toLowerCase();
    if (["success", "ok", "done"].includes(normalizedStatus)) {
      return { success: true, businessCode: data.status, message: data.message || data.msg || "" };
    }

    if (["fail", "failed", "error"].includes(normalizedStatus)) {
      return {
        success: false,
        businessCode: data.status,
        message: data.message || data.msg || `status=${data.status}`,
      };
    }
  }

  return { success: true, businessCode: null, message: "" };
}

function extractEid(data) {
  if (!data) {
    return null;
  }

  if (typeof data === "string") {
    return data;
  }

  if (typeof data.data === "string") {
    return data.data;
  }

  if (typeof data.eid === "string") {
    return data.eid;
  }

  if (typeof data.data?.eid === "string") {
    return data.data.eid;
  }

  return null;
}

function getRequestDefinitions() {
  return [
    {
      key: "participants",
      name: "participants",
      buildRequest: () => ({
        method: "post",
        url: "https://mc1.feibing.tech/api/v1/participants",
        headers: BASE_HEADERS,
        data: PARTICIPANTS_BODY,
      }),
    },
    {
      key: "wxid2eid",
      name: "wxid2eid",
      buildRequest: () => ({
        method: "get",
        url: `https://vinci.feibing.tech/sc/v1/sellers/${STATIC_CONFIG.sellerId}/wxid2eid`,
        headers: BASE_HEADERS,
        params: {
          unionid: STATIC_CONFIG.unionId,
          openid: STATIC_CONFIG.openId,
        },
      }),
      onSuccess: (responseData, state) => {
        const eid = extractEid(responseData);
        if (eid) {
          state.eid = eid;
        }
      },
    },
    {
      key: "eid2uid",
      name: "eid2uid",
      buildRequest: (state) => ({
        method: "get",
        url: "https://vinci.feibing.tech/sc/v1/ydd/eid2uid",
        headers: BASE_HEADERS,
        params: {
          eid: state.eid,
        },
      }),
    },
  ];
}

function createCouponSendRequestDefinitions(sendItems) {
  return sendItems.map((item, index) => {
    const orderRequestSn = createOrderRequestSn();

    return {
      key: `couponSend${index + 1}`,
      name: `coupon/send#${index + 1} (ticket_id=${item.ticket_id})`,
      resultExtras: {
        sendIndex: index + 1,
        ticket_id: item.ticket_id,
        total: item.total,
        order_sn: orderRequestSn,
        request_sn: orderRequestSn,
      },
      buildRequest: () => ({
        method: "post",
        url: "https://vinci.feibing.tech/sc/v1/ydd/coupon/send",
        headers: BASE_HEADERS,
        data: {
          ticket_id: item.ticket_id,
          openId: STATIC_CONFIG.couponOpenId,
          total: item.total,
          order_sn: orderRequestSn,
          request_sn: orderRequestSn,
        },
      }),
    };
  });
}

function createRequestError(requestName, requestResult, message) {
  const error = new Error(message);
  error.requestName = requestName;
  error.requestResult = requestResult;
  return error;
}

async function executeRequest(requestDefinition, state, roundNumber, timeout, logFile) {
  const startedAt = new Date().toISOString();
  const startedMs = Date.now();
  const requestConfig = requestDefinition.buildRequest(state);
  const resultExtras = requestDefinition.resultExtras || {};

  appendLog(
    logFile,
    `[第 ${roundNumber} 次] 开始请求 ${requestDefinition.name} (${requestConfig.method.toUpperCase()} ${requestConfig.url})`
  );

  let response;
  try {
    response = await axios({
      timeout,
      decompress: true,
      validateStatus: () => true,
      ...requestConfig,
    });
  } catch (error) {
    const requestResult = {
      requestKey: requestDefinition.key,
      requestName: requestDefinition.name,
      method: requestConfig.method.toUpperCase(),
      url: requestConfig.url,
      startedAt,
      finishedAt: new Date().toISOString(),
      durationMs: Date.now() - startedMs,
      httpStatus: error.response?.status ?? null,
      businessCode: null,
      success: false,
      responsePreview: safePreview(error.response?.data),
      errorMessage: error.message,
      ...resultExtras,
    };

    throw createRequestError(
      requestDefinition.name,
      requestResult,
      `${requestDefinition.name} 请求失败: ${error.message}`
    );
  }

  const business = evaluateBusinessResult(response.data);
  const success = response.status >= 200 && response.status < 300 && business.success;
  const requestResult = {
    requestKey: requestDefinition.key,
    requestName: requestDefinition.name,
    method: requestConfig.method.toUpperCase(),
    url: requestConfig.url,
    startedAt,
    finishedAt: new Date().toISOString(),
    durationMs: Date.now() - startedMs,
    httpStatus: response.status,
    businessCode: business.businessCode,
    success,
    responsePreview: safePreview(response.data),
    errorMessage: success ? "" : business.message || `HTTP ${response.status}`,
    ...resultExtras,
  };

  if (!success) {
    throw createRequestError(
      requestDefinition.name,
      requestResult,
      `${requestDefinition.name} 接口异常: ${requestResult.errorMessage}`
    );
  }

  appendLog(
    logFile,
    `[第 ${roundNumber} 次] ${requestDefinition.name} 请求成功，HTTP ${response.status}，耗时 ${requestResult.durationMs}ms`
  );

  if (typeof requestDefinition.onSuccess === "function") {
    requestDefinition.onSuccess(response.data, state);
  }

  return requestResult;
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const files = createRunFiles();
  const requestDefinitions = getRequestDefinitions();
  const sendItems = normalizeSendItems(SEND_ITEMS);

  const summary = {
    startedAt: new Date().toISOString(),
    finishedAt: null,
    loopCount: options.count,
    delayMs: options.delayMs,
    timeout: options.timeout,
    sendItems,
    sendConfigSource: "in-code constant SEND_ITEMS",
    stopped: false,
    stoppedAtRound: null,
    success: false,
    logFile: files.logFile,
    errorFile: files.errorFile,
    summaryFile: files.summaryFile,
    results: [],
    lastError: null,
  };

  appendLog(files.logFile, `开始执行发券链路，计划循环 ${options.count} 次。`);
  appendLog(
    files.logFile,
    `send 配置来源: 代码内 SEND_ITEMS 常量，本轮每次会调用 ${sendItems.length} 次 send 接口。`
  );
  writeSummary(files.summaryFile, summary);

  for (let roundNumber = 1; roundNumber <= options.count; roundNumber += 1) {
    const state = {
      eid: STATIC_CONFIG.initialEid,
    };
    const sendRequestDefinitions = createCouponSendRequestDefinitions(sendItems);

    const roundResult = {
      roundNumber,
      startedAt: new Date().toISOString(),
      finishedAt: null,
      success: false,
      sendCount: sendItems.length,
      sendItems,
      requests: [],
      error: null,
    };

    appendLog(
      files.logFile,
      `================ 第 ${roundNumber}/${options.count} 次循环开始，本轮计划调用 ${sendItems.length} 次 send 接口 ================`
    );

    let roundFailed = false;

    for (const requestDefinition of [...requestDefinitions, ...sendRequestDefinitions]) {
      try {
        const requestResult = await executeRequest(
          requestDefinition,
          state,
          roundNumber,
          options.timeout,
          files.logFile
        );
        roundResult.requests.push(requestResult);
      } catch (error) {
        if (error.requestResult) {
          roundResult.requests.push(error.requestResult);
        }

        roundFailed = true;
        roundResult.error = {
          requestName: error.requestName || requestDefinition.name,
          message: error.message,
        };

        if (error.requestResult) {
          appendErrorLog(files.errorFile, [
            `第 ${roundNumber} 次循环失败`,
            `异常接口: ${roundResult.error.requestName}`,
            `异常信息: ${roundResult.error.message}`,
            `HTTP状态: ${error.requestResult.httpStatus ?? "无"}`,
            `业务状态: ${error.requestResult.businessCode ?? "无"}`,
            `响应摘要: ${error.requestResult.responsePreview || "无"}`,
          ]);
        } else {
          appendErrorLog(files.errorFile, [
            `第 ${roundNumber} 次循环失败`,
            `异常接口: ${roundResult.error.requestName}`,
            `异常信息: ${roundResult.error.message}`,
          ]);
        }

        appendLog(
          files.logFile,
          `[第 ${roundNumber} 次] 执行已停止，异常接口: ${roundResult.error.requestName}，异常信息: ${roundResult.error.message}`
        );
        break;
      }
    }

    roundResult.finishedAt = new Date().toISOString();
    roundResult.success = !roundFailed;
    summary.results.push(roundResult);
    writeSummary(files.summaryFile, summary);

    if (roundFailed) {
      summary.stopped = true;
      summary.stoppedAtRound = roundNumber;
      summary.lastError = roundResult.error;
      summary.finishedAt = new Date().toISOString();
      summary.success = false;
      writeSummary(files.summaryFile, summary);
      appendLog(files.logFile, `任务失败，已在第 ${roundNumber} 次循环停止。`);
      throw new Error(
        `第 ${roundNumber} 次循环失败，异常接口: ${roundResult.error.requestName}，异常信息: ${roundResult.error.message}`
      );
    }

    appendLog(
      files.logFile,
      `[第 ${roundNumber} 次] 本轮接口全部成功，其中 send 接口执行 ${sendItems.length} 次。`
    );

    if (roundNumber < options.count && options.delayMs > 0) {
      appendLog(files.logFile, `等待 ${options.delayMs}ms 后进入下一轮。`);
      await sleep(options.delayMs);
    }
  }

  summary.finishedAt = new Date().toISOString();
  summary.success = true;
  writeSummary(files.summaryFile, summary);
  appendLog(files.logFile, `任务完成，${options.count} 次循环全部成功。`);
}

main().catch((error) => {
  console.error(`脚本执行失败: ${error.message}`);
  process.exit(1);
});
