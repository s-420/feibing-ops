/**
 * @对象    订单
 * @动作    批量写
 * @风险    高
 * @来源    store-data-extractor/PiLiangTuiKuan.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const axios = require("axios");

/**
 * 顺序请求退货提醒自动申请接口（一个完成再执行下一个）
 * @param {string[]} orderIds - 订单ID一维数组
 */
async function sequentialRequestReturnRemind(orderIds) {
  // 基础配置（完全保留原 curl 的请求头和URL）
  const baseUrl =
    "https://vinci-api.feibing.tech/sc/v1/sellers/wpV3cNDAAA953oOJX_L8bG7c7W0ip4Xw/contents/returnRemind/autoApplys";
  const headers = {
    accept: "*/*",
    "accept-language": "zh-CN,zh;q=0.9",
    authorization:
      "bearer __VINCI_TOKEN__",
    cookie:
      "x-token=__VINCI_TOKEN__",
    origin: "https://connect.feibing.tech",
    priority: "u=1, i",
    referer: "https://connect.feibing.tech/",
    "sec-ch-ua":
      '"Chromium";v="142", "Google Chrome";v="142", "Not_A Brand";v="99"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"macOS"',
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-site",
    "user-agent":
      "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36",
  };

  // 输入校验
  if (!Array.isArray(orderIds) || orderIds.length === 0) {
    console.error("❌ 错误：orderIds 必须是非空的一维字符串数组");
    return;
  }

  console.log(
    `🚀 开始处理 ${orderIds.length} 个订单，顺序执行（一个完成再执行下一个）\n`
  );

  // 核心：for...of + await 实现顺序执行
  for (let i = 0; i < orderIds.length; i++) {
    const orderId = orderIds[i];
    const currentIndex = i + 1;
    const totalCount = orderIds.length;

    // 过滤无效订单ID
    if (!orderId || typeof orderId !== "string") {
      console.log(
        `⚠️ [${currentIndex}/${totalCount}] 跳过无效订单ID：${orderId}（必须是非空字符串）`
      );
      continue;
    }

    try {
      console.log(
        `🔄 [${currentIndex}/${totalCount}] 正在请求订单：${orderId}`
      );

      // 发起请求（await 会等待请求完成才继续）
      const response = await axios.get(baseUrl, {
        headers,
        params: { orderIds: orderId }, // 单个请求仅传一个订单ID
        timeout: 15000, // 超时时间 15 秒（可调整）
      });

      // 按状态码 200 判断成功
      const isSuccess = response.status === 200;
      if (isSuccess) {
        console.log(
          `✅ [${currentIndex}/${totalCount}] 订单 ${orderId} - 请求成功（状态码：${response.status}）\n`
        );
      } else {
        console.log(
          `❌ [${currentIndex}/${totalCount}] 订单 ${orderId} - 请求失败（状态码：${response.status}）\n`
        );
      }
    } catch (error) {
      // 捕获所有错误（网络错误、超时、接口报错等）
      const errorReason = error.response
        ? `状态码：${error.response.status}`
        : error.message === "timeout of 15000ms exceeded"
        ? "请求超时"
        : error.message;

      console.log(
        `❌ [${currentIndex}/${totalCount}] 订单 ${orderId} - 请求失败（原因：${errorReason}）\n`
      );
    }
  }

  console.log("🎉 所有订单请求处理完毕！");
}

// -------------------------- 配置你的订单ID数组 --------------------------
const orderIds = [
  "32146748536949632",
  "32146729340537728",
  "32146728306903936",
  "32146723984280448",
  "32146678119566208",
  "32146677257505664",
  "32146676447218560",
  "32146673861823360",
  "32146672535636864",
  "32146673491020672",
  "32146670793296768",
  "32146670331005824",
  "32146668940200832",
  "32146667089333120",
  "32146665592490880",
  "32146664407206784",
  "32146663032130432",
  "32146662169938816",
  "32146658904542080",
  "32146654301424512",
  "32146649694505856",
  "32146644352797568",
  "32145054911859584",
  "32145026149550976"
];

// 执行顺序请求
sequentialRequestReturnRemind(orderIds);
