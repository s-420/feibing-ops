/**
 * @对象    门店
 * @动作    导出
 * @风险    低
 * @来源    store-data-extractor/DaoChuQingDaoPiJiuShuJuHuaNan.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

// 通用Curl解析版：导出数据（Img数组拆分→动态追加列）
const axios = require('axios');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const https = require('https');
const querystring = require('querystring');


// ===================== 可配置项：仅需修改此处Curl字符串 =====================
const curlStr = `curl 'https://qdbmgm.tsingtao.com.cn/Display/FindWhiteListDisplayReviewList' \
  -H 'Accept: application/json, text/javascript, */*; q=0.01' \
  -H 'Accept-Language: zh-CN,zh;q=0.9' \
  -H 'Connection: keep-alive' \
  -H 'Content-Type: application/x-www-form-urlencoded; charset=UTF-8' \
  -b $'__RequestVerificationToken=0VRMoARWyCi5GCoMq03kPtjTC4YiCA8fu4PDPdTsGij8F2faSmapWtXorDvBk9Nd8wkH5saI3NgN69sFJs-ITpdxIbA5v1ap7mA95G6bmMg1; csd_Id=182; csd_BCPId=39; csd_BSBId=5; csd_BesName=é\u009d’å²›å\u008dŽå\u008d—-å›¾ç‰‡å®¡æ ¸-1; csd_BId=-1; csd_BesCode=cbes_1kd8ko; csd_BComId=1; csd_IsValidate=1; csd_BesIcon=/MyStyle/images/icon/coin.jpg; csd_SpecialStatus=0; csd_ShopOnly=0; csd_CBDId=0; csd_CBDName=; useTagInfo={"bsinfo":"[]","total":"0"}; csdauthdatas=%5B%7B%22id%22%3A381%2C%22url%22%3A%22/Display/WhiteListDisplayReview%22%2C%22text%22%3A%22%u534E%u5357%u9648%u5217%u5BA1%u6838%22%2C%22icon%22%3A%22fa%20fa-file-text%22%2C%22menus%22%3A%22%22%7D%5D' \
  -H 'Origin: https://qdbmgm.tsingtao.com.cn' \
  -H 'Referer: https://qdbmgm.tsingtao.com.cn/Display/WhiteListDisplayReview' \
  -H 'Sec-Fetch-Dest: empty' \
  -H 'Sec-Fetch-Mode: cors' \
  -H 'Sec-Fetch-Site: same-origin' \
  -H 'User-Agent: Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36' \
  -H 'X-Requested-With: XMLHttpRequest' \
  -H 'sec-ch-ua: "Google Chrome";v="141", "Not?A_Brand";v="8", "Chromium";v="141"' \
  -H 'sec-ch-ua-mobile: ?0' \
  -H 'sec-ch-ua-platform: "macOS"' \
  --data-raw 'status=2%2C4&crid=&cfid=&shopname=&nickname=&hntname=&btime=2025%2F11%2F19+00%3A00%3A00&etime=2025%2F11%2F19+23%3A59%3A59.998&rewardstatus=3&bsid=143726&pageindex=1&pagesize=100'`;

  // 替换status状态： status=1%2C2%2C4%2C3   导出的表格需要根据id去重复
// ===================== 固定配置（无需修改） =====================
const outputDir = path.join(__dirname, '../output'); // 输出目录
const IMG_BASE_URL = 'https://qdbasepic.tsingtao.com.cn/HNChannelWhiteDisplayImg/'; // 华南的
// const IMG_BASE_URL = 'https://qdpic.tsingtao.com.cn/HNTaskImg/'; // 江苏的
let requestConfig = {}; // 解析后的请求配置

// ===================== 工具函数：创建输出目录 =====================
const createOutputDir = () => {
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
    console.log(`✅ 创建输出目录：${outputDir}`);
  }
};

// ===================== 工具函数：编码Cookie中的非ASCII字符 =====================
const encodeCookieValue = (value) => {
  return encodeURIComponent(value).replace(/%20/g, '+');
};

// ===================== 核心工具：解析Curl字符串 =====================
const parseCurl = (curl) => {
  try {
    console.log('📦 开始解析Curl配置...');
    const result = {
      url: '',
      method: 'post',
      headers: {},
      data: new URLSearchParams(),
      httpsAgent: new https.Agent({ rejectUnauthorized: false })
    };

    // 1. 提取URL（匹配单引号/双引号包裹的URL）
    const urlMatch = curl.match(/curl\s+['"](.+?)['"]/);
    if (!urlMatch || !urlMatch[1]) throw new Error('未找到Curl中的URL');
    result.url = urlMatch[1].trim();
    console.log(`🔗 解析URL：${result.url}`);

    // 2. 提取所有-H请求头
    const headerMatches = curl.match(/-H\s+['"](.+?)['"]/g);
    if (headerMatches) {
      headerMatches.forEach(headerStr => {
        const header = headerStr.replace(/-H\s+['"]/, '').replace(/['"]$/, '').trim();
        const [key, value] = header.split(/:\s+/, 2);
        if (key && value) {
          result.headers[key.trim()] = value.trim();
        }
      });
    }

    // 3. 提取-b Cookie（合并到headers）
    const cookieMatch = curl.match(/-b\s+\$?'(.+?)'/);
    if (cookieMatch && cookieMatch[1]) {
      let cookie = cookieMatch[1].trim();
      // 编码Cookie中的非ASCII字符
      const cookieParts = cookie.split('; ');
      const encodedCookieParts = cookieParts.map(part => {
        const [key, value] = part.split('=', 2);
        if (key && value && /[^\x00-\x7F]/.test(value)) { // 检测非ASCII字符
          return `${key}=${encodeCookieValue(value)}`;
        }
        return part;
      });
      result.headers.Cookie = encodedCookieParts.join('; ');
    }

    // 4. 提取--data-raw表单数据
    const dataMatch = curl.match(/--data-raw\s+['"](.+?)['"]/);
    if (dataMatch && dataMatch[1]) {
      const dataStr = decodeURIComponent(dataMatch[1].trim());
      const dataObj = querystring.parse(dataStr);
      Object.entries(dataObj).forEach(([key, value]) => {
        result.data.set(key, value.toString());
      });
    }

    console.log(`✅ Curl解析完成，提取配置：`);
    console.log(`- 请求方法：${result.method}`);
    console.log(`- 请求头数量：${Object.keys(result.headers).length} 个`);
    console.log(`- 表单参数：${Array.from(result.data.keys()).join(', ')}`);
    return result;
  } catch (error) {
    console.error('❌ Curl解析失败：', error.message);
    process.exit(1);
  }
};

// ===================== 工具函数：格式化日期 =====================
const formatDate = (dateString) => {
  if (!dateString || !dateString.includes('/Date(')) return dateString;
  const timestamp = parseInt(dateString.match(/\d+/)[0]);
  const date = new Date(timestamp);
  return date.toISOString().slice(0, 19).replace('T', ' ');
};

// ===================== 工具函数：处理数据（Img数组拆分+拼接域名） =====================
const processData = (item) => {
  const processed = { ...item };
  
  // 1. 格式化日期字段
  Object.keys(processed).forEach(key => {
    if (key.includes('Time') && typeof processed[key] === 'string') {
      processed[key] = formatDate(processed[key]);
    }
  });

  // 2. 处理TaskImg（保持原有逻辑+转string）
  if (processed.TaskImg) {
    processed.TaskImg = String(`${IMG_BASE_URL}${processed.TaskImg}`);
  } else {
    processed.TaskImg = '';
  }

  // 3. 核心处理：Img字段（逗号拆分→数组→拼接域名→清空原字段）
  processed.imgUrls = []; // 存储拆分后拼接好的URL数组
  if (processed.Img) {
    // 转为string→按逗号拆分→过滤空元素→清理无效字符
    const imgArray = String(processed.Img)
      .split(',')
      .map(img => img.trim())
      .filter(img => img); // 过滤空字符串
    
    // 每个元素拼接域名→转为string
    processed.imgUrls = imgArray.map(img => String(`${IMG_BASE_URL}${img}`));
  }
  processed.Img = ''; // 清空原Img列，不再存放数据

  return processed;
};

// ===================== 主函数 =====================
const main = async () => {
  try {
    // 1. 初始化（创建目录 + 解析Curl）
    createOutputDir();
    requestConfig = parseCurl(curlStr);

    // 2. 初始化Excel + 循环请求数据
    let pageIndex = 1;
    let allItems = [];
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('数据列表');

    // 循环请求分页数据
    while (true) {
      console.log(`\n📥 正在获取第 ${pageIndex} 页数据...`);
      
      // 更新当前页码（适配Curl中的pageindex参数）
      requestConfig.data.set('pageindex', pageIndex.toString());
      
      // 发送请求
      const response = await axios(requestConfig);
      const data = response.data;
      
      // 无数据则停止
      if (!data || !data.items || data.items.length === 0) {
        console.log('📭 没有更多数据，停止循环');
        break;
      }
      
      // 处理数据（Img数组拆分+拼接域名）
      const processedItems = data.items.map(item => processData(item));
      allItems = [...allItems, ...processedItems];
      console.log(`✅ 第 ${pageIndex} 页处理完成，共 ${processedItems.length} 条`);
      
      pageIndex++;
    }

    // 3. 生成Excel（动态追加Img列）
    if (allItems.length > 0) {
      // 3.1 计算最大Img数组长度（用于动态生成列）
      const maxImgCount = allItems.reduce((max, item) => {
        return Math.max(max, item.imgUrls.length);
      }, 0);
      console.log(`📊 检测到最大Img数组长度：${maxImgCount}，将追加 ${maxImgCount} 列（Img1~Img${maxImgCount}）`);

      // 3.2 构建表头：原始字段 + 动态Img列（Img1、Img2...）
      const rawHeaders = Object.keys(allItems[0]).filter(key => key !== 'imgUrls'); // 排除临时存储的imgUrls
      const imgHeaders = maxImgCount > 0 
        ? Array.from({ length: maxImgCount }, (_, i) => `Img${i + 1}`) 
        : [];
      const headers = [...rawHeaders, ...imgHeaders];

      // 3.3 添加表头样式
      const headerRow = worksheet.addRow(headers);
      headerRow.eachCell(cell => {
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E6F3FF' } };
        cell.alignment.wrapText = true;
      });

      // 3.4 添加数据行（原始数据 + 动态Img列数据）
      allItems.forEach((item, index) => {
        const rowIndex = index + 2;
        // 原始数据（排除imgUrls，原Img列已清空）
        const rawRowData = rawHeaders.map(header => {
          return item[header] === null || item[header] === undefined ? '' : item[header];
        });
        // Img列数据：不足maxImgCount时填空字符串
        const imgRowData = imgHeaders.map((_, i) => {
          return item.imgUrls[i] || '';
        });
        // 合并数据行
        const rowData = [...rawRowData, ...imgRowData];
        const dataRow = worksheet.addRow(rowData);
        
        // 数据行样式
        dataRow.eachCell(cell => {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          if (cell.value && typeof cell.value === 'string' && cell.value.length > 20) {
            cell.alignment.wrapText = true;
          }
        });
      });

      // 3.5 列宽优化
      headers.forEach((header, index) => {
        const column = worksheet.getColumn(index + 1);
        let columnWidth = header.length + 2;
        
        // 计算最大内容长度
        allItems.forEach(item => {
          let value = '';
          // 原始字段值
          if (rawHeaders.includes(header)) {
            value = item[header] || '';
          } 
          // Img列值
          else if (header.startsWith('Img')) {
            const imgIndex = parseInt(header.replace('Img', '')) - 1;
            value = item.imgUrls[imgIndex] || '';
          }
          
          if (value && typeof value === 'string') {
            columnWidth = Math.min(Math.max(columnWidth, value.length + 2), 60);
          }
        });
        
        // 特殊字段宽度
        if (header.startsWith('Img')) columnWidth = 60; // 新增的Img列
        else if (header === 'Address' || header === 'HNTName' || header === 'ShopName') columnWidth = 40;
        else if (header.includes('Time') || header === 'TaskImg') columnWidth = 22;
        
        column.width = Math.max(columnWidth, 12);
      });

      // 3.6 行高优化
      worksheet.getRow(1).height = 30;
      for (let i = 2; i <= allItems.length + 1; i++) {
        worksheet.getRow(i).height = 35;
      }

      // 3.7 保存文件
      const fileName = `数据导出_含动态Img列_${new Date().toLocaleDateString().replace(/\//g, '-')}.xlsx`;
      const filePath = path.join(outputDir, fileName);
      await workbook.xlsx.writeFile(filePath);
      
      console.log(`\n📁 导出完成：${filePath}`);
      console.log(`📊 统计：${allItems.length} 条记录，${headers.length} 个字段（含 ${maxImgCount} 个Img列）`);
      console.log(`✅ 原Img列已清空，拆分后的数据已追加到 Img1~Img${maxImgCount} 列`);
    } else {
      console.log('📭 未获取到任何数据');
    }
  } catch (error) {
    console.error('\n❌ 执行失败：', error.message);
    if (error.response) {
      console.error('🔍 响应状态：', error.response.status);
      console.error('🔍 响应数据：', error.response.data);
    } else if (error.request) {
      console.error('🔍 请求未收到响应（可能是URL/Cookie错误）');
    }
  }
};

// 启动脚本
main();