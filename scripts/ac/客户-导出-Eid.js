/**
 * @对象    客户
 * @动作    导出
 * @风险    低
 * @来源    store-data-extractor/DaoChuKeHuEidXinXi.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const axios = require('axios');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// API请求配置
const API_URL = 'https://scrm.feibing.tech/wework-scrm/customer/list';
const headers = {
  'accept': 'application/json',
  'accept-language': 'zh-CN,zh;q=0.9',
  'content-type': 'application/json;charset=UTF-8',
  'origin': 'https://scrm.feibing.tech',
  'priority': 'u=1, i',
  'referer': 'https://scrm.feibing.tech/',
  'sec-ch-ua': '"Chromium";v="142", "Google Chrome";v="142", "Not_A Brand";v="99"',
  'sec-ch-ua-mobile': '?0',
  'sec-ch-ua-platform': '"macOS"',
  'sec-fetch-dest': 'empty',
  'sec-fetch-mode': 'cors',
  'sec-fetch-site': 'same-origin',
  'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36',
  'Cookie': '_bl_uid=aa702ba8-862d-4f10-97ea-08f95927e9ed; x-token=__VINCI_TOKEN__'
};

// 请求基础数据
const baseRequestData = {
  "computeType": 1,
  "pageSize": 100,
  "followerUserId": "woV3cNDAAAgAxbCKkXQbu1t1_kvAmROA",
  "tagIdList": [],
  "addTime": {
    "startTime": 1761926400000,
    "endTime": 1764172799999
  }
};

// 存储所有客户数据
const allCustomers = [];

/**
 * 递归获取所有分页的客户数据
 * @param {number} page - 当前页码
 */
async function fetchCustomerData(page = 1) {
  try {
    console.log(`正在获取第 ${page} 页数据...`);

    // 发送请求
    const response = await axios.post(
      API_URL,
      { ...baseRequestData, page },
      { headers }
    );

    // 检查响应状态
    if (response.data.code !== 200) {
      console.error(`请求失败，错误码: ${response.data.code}`);
      return;
    }

    const { list, totalPage, currPage } = response.data.data;

    // 提取需要的字段
    const formattedData = list.map(customer => ({
      name: customer.name,
      externalUserId: customer.externalUserId
    }));

    // 添加到总数据中
    allCustomers.push(...formattedData);
    console.log(`第 ${currPage} 页获取完成，共 ${formattedData.length} 条数据`);

    // 如果还有下一页，继续请求
    if (currPage < totalPage) {
      await fetchCustomerData(currPage + 1);
    } else {
      console.log(`所有数据获取完成，共 ${allCustomers.length} 条记录`);
      // 导出到Excel
      exportToExcel(allCustomers);
    }
  } catch (error) {
    console.error('请求出错:', error.message);
    if (error.response) {
      console.error('响应状态:', error.response.status);
      console.error('响应数据:', error.response.data);
    }
  }
}

/**
 * 将数据导出为Excel文件
 * @param {Array} data - 要导出的数据
 */
function exportToExcel(data) {
  try {
    // 创建工作表
    const worksheet = XLSX.utils.json_to_sheet(data);

    // 创建工作簿并添加工作表
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '客户列表');

    // 确保output文件夹存在
    const outputDir = path.join(__dirname, '../output');
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    // 保存文件
    const filePath = path.join(outputDir, '客户Eid.xlsx');
    XLSX.writeFile(workbook, filePath);

    console.log(`Excel文件已成功导出至: ${filePath}`);
  } catch (error) {
    console.error('导出Excel失败:', error.message);
  }
}

// 开始执行
console.log('开始获取客户数据...');
fetchCustomerData();