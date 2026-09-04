/**
 * @对象    客户
 * @动作    导出
 * @风险    低
 * @来源    store-data-extractor/DaoChuKeHuLieBiao.js
 * @脱敏    硬编码 token 已替换为 __VINCI_TOKEN__，运行前请改为从 .env 读取 VINCI_AUTHORIZATION
 */

const axios = require('axios');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 配置信息
const config = {
  // 请求的基础URL
  baseUrl: 'https://scrm.feibing.tech/wework-scrm/customer/list',
  // 每页大小（改为100）
  pageSize: 100,
  // 请求头
  headers: {
    'accept': 'application/json',
    'accept-language': 'zh-CN,zh;q=0.9',
    'content-type': 'application/json;charset=UTF-8',
    'origin': 'https://scrm.feibing.tech',
    'priority': 'u=1, i',
    'referer': 'https://scrm.feibing.tech/',
    'sec-ch-ua': '"Not:A-Brand";v="99", "Google Chrome";v="145", "Chromium";v="145"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"macOS"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/145.0.0.0 Safari/537.36',
    'Cookie': '_bl_uid=ab3b0489-1a0b-40e0-b8a3-5f5685e318fb; x-token=__VINCI_TOKEN__'
  },
  // 时间范围
  addTime: {
    startTime: 1772208000000,
    endTime: 1773849599999
  },
  // 输出目录
  outputDir: './exported_excels',
  // 汇总文件名
  summaryFileName: '客户数据汇总.xlsx',

  // 并发配置
  storeConcurrency: 5, // 门店并发数
  pageConcurrency: 4,  // 单门店分页并发数

  // followerUserList 保持你原来的完整列表
  followerUserList: [
    {
      "name": "南桥百联店",
      "id": "woV3cNDAAALsdy9AJWo0z8vRBGL6uGEQ"
    },
    {
      "name": "浦江满天星店",
      "id": "woV3cNDAAAG2qPHKniC4nOSV-uWbYLVA"
    },
    {
      "name": "福州路店",
      "id": "woV3cNDAAAgE5H7KJDQouF8GgWNtt7SQ"
    },
    {
      "name": "松兰路",
      "id": "woV3cNDAAAoVVM7HHhmbigmDPupFl23g"
    },
    {
      "name": "北中路店",
      "id": "woV3cNDAAA7V66hAXbaD6yAzeb_X1gcA"
    },
    {
      "name": "昌里上南店",
      "id": "woV3cNDAAAz3_iaUDFH23s46Dlzd3-2Q"
    },
    {
      "name": "万科天空之城店",
      "id": "woV3cNDAAAFIiAS8O3ce1_n-APVBe12g"
    },
    {
      "name": "开新里店",
      "id": "woV3cNDAAA9LPr3G7gcr1lqbQoZanWBA"
    },
    {
      "name": "邯郸店",
      "id": "woV3cNDAAArNW1yxGqabXGL2ta4P6yJA"
    },
    {
      "name": "马陆地铁店",
      "id": "woV3cNDAAA4LPsn1oi13Zp84mi2qSlcg"
    },
    {
      "name": "喜泉店",
      "id": "woV3cNDAAAvp7ClhX5W_MN5KUeYO0Uyw"
    },
    {
      "name": "桃浦智创城店",
      "id": "woV3cNDAAADDZVkIUEfcArBmpxXUYkfA"
    },
    {
      "name": "荣科店",
      "id": "woV3cNDAAAN7fgSBy2PoPsYhCJIIjwug"
    },
    {
      "name": "南翔印象城店",
      "id": "woV3cNDAAAvhRFPmCSbyfVxupg_Z4CeQ"
    },
    {
      "name": "老西门店",
      "id": "woV3cNDAAA31yeCortfmJLWXXrDz3hTg"
    },
    {
      "name": "奉贤海泉店",
      "id": "woV3cNDAAA1FSKzgK7dQ6i0gDgIUN-Cw"
    },
    {
      "name": "1点点外滩中心店",
      "id": "woV3cNDAAAppfTQFWGLStn7YAnAEUs-Q"
    },
    {
      "name": "MOHO店",
      "id": "woV3cNDAAAxSxVHom8RcP8y4cx5WzUWw"
    },
    {
      "name": "交大番禺店",
      "id": "woV3cNDAAAOh-AdwmV_NNctpeYh56EyQ"
    },
    {
      "name": "上海中心店",
      "id": "woV3cNDAAA9mDuSP-WNXlHquWQcJH0mg"
    },
    {
      "name": "特斯拉工厂店",
      "id": "woV3cNDAAA1HRiKSpQfNkf6JjSugvpnQ"
    },
    {
      "name": "安亭地铁店",
      "id": "woV3cNDAAAAs6oYNs_0o0X2LdWWH-wVg"
    },
    {
      "name": "奉贤大润发店",
      "id": "woV3cNDAAAEff_oPmr2BQ8Jqz7fOt5Pg"
    },
    {
      "name": "森兰广场店",
      "id": "woV3cNDAAAAY7IT1QHNSsh5NwDylIjyw"
    },
    {
      "name": "龙湖金汇店",
      "id": "woV3cNDAAATjOuJ5DLN6t842wzmlxDhQ"
    },
    {
      "name": "金桥店",
      "id": "woV3cNDAAAC_Ut-P88MWyH_Ul0DUssCQ"
    },
    {
      "name": "1点点1788广场店",
      "id": "woV3cNDAAALVfz7eBaumUYy6vzXjyS3w"
    },
    {
      "name": "1点点南丰城店",
      "id": "woV3cNDAAA76ZNT0VKhG5koUX58YNMXA"
    },
    {
      "name": "百盛优客店",
      "id": "woV3cNDAAAjz_WHID_8itqJeS0rPBP7g"
    },
    {
      "name": "宝杨宝龙店",
      "id": "woV3cNDAAAHlOmLjVJMCrGRpl7ggNUNw"
    },
    {
      "name": "成山巴黎春天店",
      "id": "woV3cNDAAARFak59DX57vQMijTEhp7wg"
    },
    {
      "name": "大华乐坊店",
      "id": "woV3cNDAAAmcLExCu3x-9JjWjr1eAsLQ"
    },
    {
      "name": "东方万国店",
      "id": "woV3cNDAAAQFILWheUZmjZ3XYTouk6pA"
    },
    {
      "name": "恒越荣欣店",
      "id": "woV3cNDAAAGX-qhHbwoFSccPSbnVfGFA"
    },
    {
      "name": "虹桥天地店",
      "id": "woV3cNDAAAPcyl6RXYXh6f2WQ02dbLaw"
    },
    {
      "name": "花都荟店",
      "id": "woV3cNDAAAfpN-67kWl_DVpt-FtWKNYg"
    },
    {
      "name": "龙湖闵行天街店",
      "id": "woV3cNDAAA8-Fgc7pPtl6O8dKaDctWTg"
    },
    {
      "name": "马桥万达店",
      "id": "woV3cNDAAAsIXXZZBONqB6v0Q1rZvJcQ"
    },
    {
      "name": "闵行爱琴海店",
      "id": "woV3cNDAAAPJhgT75Lvs83UvMzs2IY5w"
    },
    {
      "name": "浦江万达店",
      "id": "woV3cNDAAAShNW9YDbue7_TGKF00GTfQ"
    },
    {
      "name": "莘庄维璟印象城店",
      "id": "woV3cNDAAAHpr9IMYMGzUY9Iw_414sxA"
    },
    {
      "name": "孙桥店",
      "id": "woV3cNDAAArZrHFn_69WfCCNAs7PUUfw"
    },
    {
      "name": "新会店",
      "id": "woV3cNDAAAdr4G-g21FZ8TGoSujpYOrw"
    },
    {
      "name": "徐汇日月光店",
      "id": "woV3cNDAAA8numKOQZ1b5J_sU8p47hwQ"
    },
    {
      "name": "绚荟城店",
      "id": "woV3cNDAAAqlJHtahWO865oXi7n9_bBA"
    },
    {
      "name": "1点点悦荟广场店",
      "id": "woV3cNDAAA9Vwd7XVMMh3x5NgSNyK6Kg"
    },
    {
      "name": "中房金谊店",
      "id": "woV3cNDAAAzBjOfWRsTFCBEMgY1DVWeQ"
    },
    {
      "name": "1点点万嘉店",
      "id": "woV3cNDAAAPqKav2kGZDbuRA0Tok1Y5Q"
    },
    {
      "name": "TPY店",
      "id": "woV3cNDAAAeYkRWMeN8GM0xtgzTFXdqg"
    },
    {
      "name": "白玉兰店",
      "id": "woV3cNDAAAMYGpad4ye5uFGtjiObtaJg"
    },
    {
      "name": "万渡汇广场店",
      "id": "woV3cNDAAAvPdj5A0LBzZejie-o3cB5A"
    },
    {
      "name": "宝山U天地店",
      "id": "woV3cNDAAAmcLExCu3x-9JjWjr1eAsLQ"
    },
    {
      "name": "陈彷公路店",
      "id": "woV3cNDAAAgDMnj9Z6xIl8zanF2vtdeg"
    },
    {
      "name": "成山浦乐汇店",
      "id": "woV3cNDAAAtCHfzMAWAOLgJgFG9Nvu4A"
    },
    {
      "name": "崇明万达店",
      "id": "woV3cNDAAAN1iSpA0ttIyzPD6AP4-AKQ"
    },
    {
      "name": "大华第一坊店",
      "id": "woV3cNDAAAahVczFHvx_ZhFlH8qxCgjg"
    },
    {
      "name": "大宁音乐广场店",
      "id": "woV3cNDAAAbVfb7xeyVkg1ieoeaVNkBQ"
    },
    {
      "name": "东方懿德城店",
      "id": "woV3cNDAAAsiFCUkm_d73CQ0sL1qF86A"
    },
    {
      "name": "丰福店",
      "id": "woV3cNDAAAgDMnj9Z6xIl8zanF2vtdeg"
    },
    {
      "name": "合生汇店",
      "id": "woV3cNDAAAysJ6PCLl77DV4ouLIEF4fw"
    },
    {
      "name": "1点点凯德虹口商业中心店",
      "id": "woV3cNDAAAUOWSYW7Xb_mUvnb6QNg4FA"
    },
    {
      "name": "华光广场店",
      "id": "woV3cNDAAAP1vveiQOsABm__oDmVkwdQ"
    },
    {
      "name": "江湾吉买盛店",
      "id": "woV3cNDAAAepToVuARSUp5JZBM2Di-Og"
    },
    {
      "name": "金山大润发店",
      "id": "woV3cNDAAA1N0QaeHPcX4IHEiP61yEHw"
    },
    {
      "name": "京华路店",
      "id": "woV3cNDAAA-gxbeNVh-E0V_NGcFpNqew"
    },
    {
      "name": "开元地中海店",
      "id": "woV3cNDAAAtxAJ8PJONudkrzEh82W73A"
    },
    {
      "name": "康桥新田360店",
      "id": "woV3cNDAAACOZaQyo47ZvtGc-37p-LrA"
    },
    {
      "name": "李子园店",
      "id": "woV3cNDAAAyCChXAy4dALFp8rnho6Dew"
    },
    {
      "name": "临港宝龙店",
      "id": "woV3cNDAAAYYCSXVUNXvInIufuvCfIbA"
    },
    {
      "name": "龙湖北城天街店",
      "id": "woV3cNDAAAMPROl_6_BWegDIx4lTnflw"
    },
    {
      "name": "龙阳广场店",
      "id": "woV3cNDAAA7odvfijxrqaG84vfjKcqDQ"
    },
    {
      "name": "南汇大润发店",
      "id": "woV3cNDAAAdyI2-Chgg_9L67jVyV2sWQ"
    },
    {
      "name": "浦东禹悦汇店",
      "id": "woV3cNDAAASjQWoCw1iM9nQtpXXuIMbQ"
    },
    {
      "name": "浦乐生活广场店",
      "id": "woV3cNDAAAHeLFrXwxPvk-iQYSPJs6sw"
    },
    {
      "name": "七宝乐购店",
      "id": "woV3cNDAAA9mDuSP-WNXlHquWQcJH0mg"
    },
    {
      "name": "青浦宝龙店",
      "id": "woV3cNDAAAJShdIyEv-ZO5DRVTIaLbug"
    },
    {
      "name": "青浦绿地缤纷城店",
      "id": "woV3cNDAAAJUnsremTTqG7B5_sUCOfuw"
    },
    {
      "name": "青浦苏杭时代店",
      "id": "woV3cNDAAA--JHxz5g5UGm3wfIPZQ7Ug"
    },
    {
      "name": "盈港路富绅时代店",
      "id": "woV3cNDAAAOGzDtmPW3EZPtFPVrBwl2g"
    },
    {
      "name": "三林新达汇店",
      "id": "woV3cNDAAAFw4X4rEJ1pCFi7-zKDPP5g"
    },
    {
      "name": "莘庄龙之梦店",
      "id": "woV3cNDAAA8gVTyLhHquFQ4_b_ckmE3Q"
    },
    {
      "name": "世纪汇店",
      "id": "woV3cNDAAAo6uDb3H-FJMvnaSejGszDA"
    },
    {
      "name": "呈祥疏影店",
      "id": "woV3cNDAAAC_r43K3bTSKXwfMJDt2bng"
    },
    {
      "name": "万象城店",
      "id": "woV3cNDAAAvsY0daGkyD-AjANTChkqDg"
    },
    {
      "name": "五彩城店",
      "id": "woV3cNDAAAH3V3eEhHYhqF7OuSuaPRRw"
    },
    {
      "name": "西岸凤巢店",
      "id": "woV3cNDAAAOgTXrdsoQKP5QUoafLI1iQ"
    },
    {
      "name": "小世界店",
      "id": "woV3cNDAAAeYkRWMeN8GM0xtgzTFXdqg"
    },
    {
      "name": "新鞍山路店",
      "id": "woV3cNDAAA9upfNR9liSeDC_5JqrWArA"
    },
    {
      "name": "新环广场店",
      "id": "woV3cNDAAA7Qh5YWEg7YmE0oWMRYFLHw"
    },
    {
      "name": "新世界店",
      "id": "woV3cNDAAAjHaBlvOV9vwCDQxhgq6XQw"
    },
    {
      "name": "信业广场店",
      "id": "woV3cNDAAA7xkt4s47Jijd4oPtcoRkGw"
    },
    {
      "name": "亚乐城店",
      "id": "woV3cNDAAAVL07hNnME9xTVfnR4QCfxA"
    },
    {
      "name": "益江路店",
      "id": "woV3cNDAAAdKf_irF6hoFpea71sIt7hQ"
    },
    {
      "name": "长宁88店",
      "id": "woV3cNDAAAylarcyELoxauCAleb5rzwA"
    },
    {
      "name": "中骏广场店",
      "id": "woV3cNDAAAPcyl6RXYXh6f2WQ02dbLaw"
    },
    {
      "name": "昌吉店",
      "id": "woV3cNDAAAPcyl6RXYXh6f2WQ02dbLaw"
    },
    {
      "name": "大华二店",
      "id": "woV3cNDAAAmcLExCu3x-9JjWjr1eAsLQ"
    },
    {
      "name": "定西店",
      "id": "woV3cNDAAAcj5lUHhiFrYuI5AMjxF-iw"
    },
    {
      "name": "奉贤上师大店",
      "id": "woV3cNDAAAO7FibPPli0fu5yAmIne4eg"
    },
    {
      "name": "共康店",
      "id": "woV3cNDAAAjwz9i5UYpl31WrVVZt66wg"
    },
    {
      "name": "建滔广场店",
      "id": "woV3cNDAAALqsFVosVUxCjUdiTXJJebw"
    },
    {
      "name": "乐尚天地店",
      "id": "woV3cNDAAA4w9EuxYdqB1yi_-SRZOP3Q"
    },
    {
      "name": "浦江欢乐颂店",
      "id": "woV3cNDAAAE-_cmE_y3CTsrSSC4Db_-A"
    },
    {
      "name": "天钥桥店",
      "id": "woV3cNDAAAaLEl_RCstAWf_7E9PsvGhg"
    },
    {
      "name": "西藏南路店",
      "id": "woV3cNDAAAhunmOzfJLtyfQl6pJuM92w"
    },
    {
      "name": "雪松店",
      "id": "woV3cNDAAA7c_t9kDmPh68gI7pUk8Ohg"
    },
    {
      "name": "8号桥店",
      "id": "woV3cNDAAARLSkLkQhuEqWCBrJjWwNWg"
    },
    {
      "name": "金虹桥店",
      "id": "woV3cNDAAAqpNHmob-sfyxS8ZmAbU2bw"
    },
    {
      "name": "青浦富绅中心店",
      "id": "woV3cNDAAAaY-6ds0ZKLAXtkUp3aCGHg"
    },
    {
      "name": "殷高西路店",
      "id": "woV3cNDAAAix-Um8ipen0LhlCmg-c8FA"
    },
    {
      "name": "共和新路店",
      "id": "woV3cNDAAAoQEUrKrc3WEo9O8F7V6cWw"
    },
    {
      "name": "大木桥店",
      "id": "woV3cNDAAAGGNOL7iefsAZY_UkrLrlPw"
    },
    {
      "name": "德都店",
      "id": "woV3cNDAAAmchdQSNCt-I6Sdh_Bz1uKA"
    },
    {
      "name": "新水电路店",
      "id": "woV3cNDAAAix-Um8ipen0LhlCmg-c8FA"
    },
    {
      "name": "曹杨路店",
      "id": "woV3cNDAAA4LPsn1oi13Zp84mi2qSlcg"
    },
    {
      "name": "青浦万达茂广场店",
      "id": "woV3cNDAAA4LPsn1oi13Zp84mi2qSlcg"
    },
    {
      "name": "松江万达店",
      "id": "woV3cNDAAA-iQu5ujZGRRZD5r9e4puGA"
    },
    {
      "name": "韩村路店",
      "id": "woV3cNDAAA8W-iq5Qsiv9uF2qItMeexw"
    },
    {
      "name": "张江汇智店",
      "id": "woV3cNDAAAfPSfiE4r3CPSFYsfdf1E9g"
    },
    {
      "name": "真如环宇城店",
      "id": "woV3cNDAAAx4ZBWfYkMupBKOxvsI0myQ"
    },
    {
      "name": "南翔8760店",
      "id": "woV3cNDAAA68bHiytY7CgT3R2LDs-RSw"
    },
    {
      "name": "经纬汇店",
      "id": "woV3cNDAAAkK4S9rSjeHcIOcAdOp2Rrw"
    },
    {
      "name": "金山万达店",
      "id": "woV3cNDAAA6U4OyLOazfx-PgxeeuQVtQ"
    },
    {
      "name": "颛桥万达店",
      "id": "woV3cNDAAAuQRBzf4JAc9TsKYebpw71w"
    },
    {
      "name": "淮海巴黎春天店",
      "id": "woV3cNDAAA-1dsKIlxlgq3lhkXZtnnRg"
    },
    {
      "name": "曹路招商花园城店",
      "id": "woV3cNDAAAFZqq3UjhS856kPE0lJ4UfA"
    },
    {
      "name": "大同店",
      "id": "woV3cNDAAAT-nZVef9lkiQ5Wn1v0KGqw"
    },
    {
      "name": "奉贤四新店",
      "id": "woV3cNDAAAqpNHmob-sfyxS8ZmAbU2bw"
    },
    {
      "name": "杨思店(近杨思地铁站)",
      "id": "woV3cNDAAAPn4KpEocIWH4cx8wOrjLag"
    },
    {
      "name": "西郊百联",
      "id": "woV3cNDAAAVvF6K9xI5mDe6uh1YDPR9Q"
    },
    {
      "name": "崇恒新天地",
      "id": "woV3cNDAAA-D6wHJ6zKqKTvKbIqX10lw"
    },
    {
      "name": "大林店",
      "id": "woV3cNDAAAVGVYTFTrUqfCtjuNGaY9VA"
    },
    {
      "name": "徐乐北店",
      "id": "woV3cNDAAA3zQUiCfCtQjmDWiJnslzew"
    },
    {
      "name": "宝慧店",
      "id": "woV3cNDAAAqZzCZCrKQfrmXsB0Et46KQ"
    },
    {
      "name": "漫游城店",
      "id": "woV3cNDAAA2rqH3leRqHLjO8MODIOKkA"
    },
    {
      "name": "石潭街店",
      "id": "woV3cNDAAAdBWmbvg-YyC-kP3aW2Stug"
    },
    {
      "name": "银春路店",
      "id": "woV3cNDAAAuQRBzf4JAc9TsKYebpw71w"
    },
    {
      "name": "航北路店",
      "id": "woV3cNDAAAVs4F3TnxD75rhWL08SrDAg"
    },
    {
      "name": "长风VIA店",
      "id": "woV3cNDAAAVs4F3TnxD75rhWL08SrDAg"
    },
    {
      "name": "荆州店",
      "id": "woV3cNDAAAQ8J6HcMBJN9FzIrJqTlwrg"
    },
    {
      "name": "车墩金地店",
      "id": "woV3cNDAAAhydhktJHLcGOz2TW8kc7-w"
    },
    {
      "name": "南门店",
      "id": "woV3cNDAAAuJdT_TDtMlARo-3sI9CNaQ"
    },
    {
      "name": "兰溪店",
      "id": "woV3cNDAAAZ2JlDtWwUEbkFMF81fpHtg"
    },
    {
      "name": "梅川店",
      "id": "woV3cNDAAAZ2JlDtWwUEbkFMF81fpHtg"
    },
    {
      "name": "东鼎店",
      "id": "woV3cNDAAAJaj1xo5tKI467Do9S8VGYA"
    },
    {
      "name": "龙吴路店",
      "id": "woV3cNDAAApYv_Q2jHAL4VaJOb3onXtw"
    },
    {
      "name": "张堰德贤店",
      "id": "woV3cNDAAArGF9o92vWLroPsMjvjWvPQ"
    },
    {
      "name": "文涵店",
      "id": "woV3cNDAAAFdgfU3ve-X1cV-eKANYqdw"
    },
    {
      "name": "江文店",
      "id": "woV3cNDAAAMDuZTmJpRqvuzYu6HgnQZQ"
    },
    {
      "name": "民星店",
      "id": "woV3cNDAAAsGdpBGhESeMFfkzAs1TEmQ"
    },
    {
      "name": "淞南店",
      "id": "woV3cNDAAAATDa1DjvjaMscv9el0oFUg"
    },
    {
      "name": "虎林店",
      "id": "woV3cNDAAAamwEDooiKRX--eGEEGBC-Q"
    },
    {
      "name": "金航城店",
      "id": "woV3cNDAAAz3_iaUDFH23s46Dlzd3-2Q"
    },
    {
      "name": "奉贤龙湖天街店",
      "id": "woV3cNDAAASk24P0rScelUO53206DwXw"
    },
    {
      "name": "永德店",
      "id": "woV3cNDAAAkIo8vhvG_jvZNB538C9xoQ"
    },
    {
      "name": "长阳路店",
      "id": "woV3cNDAAArevEKLE9EgwELjyb_10U_Q"
    },
    {
      "name": "罗秀创邑店",
      "id": "woV3cNDAAAhIKonhp9_xvouAjirN6vaA"
    },
    {
      "name": "田林店",
      "id": "woV3cNDAAA4hCKMY2ml0BaI3iXpQfPdQ"
    },
    {
      "name": "控江店",
      "id": "woV3cNDAAAqB6E4yvRHDhVZYFWsEUVAQ"
    },
    {
      "name": "中江店",
      "id": "woV3cNDAAAlgFPa9gUt0rSuqfPGf7QCg"
    },
    {
      "name": "虹井店",
      "id": "woV3cNDAAAwoS_DYDo1STOIKFavi2mwQ"
    },
    {
      "name": "中山龙之梦店",
      "id": "woV3cNDAAAPM8bVnuf8DdZgDUzayYX-Q"
    },
    {
      "name": "沪太华阴路店",
      "id": "woV3cNDAAAnDntIUL0f-Q1d0PTaMYiZw"
    },
    {
      "name": "金悦乐方广场",
      "id": "woV3cNDAAAhIKonhp9_xvouAjirN6vaA"
    },
    {
      "name": "松汇二店",
      "id": "woV3cNDAAAB_w30Ikgu2y3T3PJWmbSTA"
    },
    {
      "name": "松汇店",
      "id": "woV3cNDAAAxjjAUtMmKcXyi1aFNJADzQ"
    },
    {
      "name": "平原店",
      "id": "woV3cNDAAAGAC06ubx2XRsqVw4yU30dQ"
    },
    {
      "name": "罗店宝龙广场店",
      "id": "woV3cNDAAAMiLOJvgyRvDcl3EG_wb-dw"
    },
    {
      "name": "堡镇镇中店",
      "id": "woV3cNDAAAdhyotSCAPwaWt6r3TnrsxA"
    },
    {
      "name": "育秀店",
      "id": "woV3cNDAAAd3L1kivYW1mBg29oqhtTVA"
    },
    {
      "name": "威宁店",
      "id": "woV3cNDAAAT-nZVef9lkiQ5Wn1v0KGqw"
    },
    {
      "name": "环镇北店",
      "id": "woV3cNDAAADqt-6Ttmg-DADvjMfXempQ"
    },
    {
      "name": "浦三店",
      "id": "woV3cNDAAAYn5GVoKWFAA-OJjGvScJOQ"
    },
    {
      "name": "文峰广场店",
      "id": "woV3cNDAAAxbqi1dBKG3EMlV0rjg-g0w"
    },
    {
      "name": "新大陆二店",
      "id": "woV3cNDAAAPds1E2QbyuI8JA_bZQ_l3Q"
    },
    {
      "name": "航头店",
      "id": "woV3cNDAAAd3L1kivYW1mBg29oqhtTVA"
    },
    {
      "name": "文汇店",
      "id": "woV3cNDAAAIrIsQZPeN2-h_eJEiQwBPg"
    },
    {
      "name": "复旦软件园店",
      "id": "woV3cNDAAAJqkpDIJRyYJW6ffv3Zv3hA"
    },
    {
      "name": "川沙路店",
      "id": "woV3cNDAAAqnrSG8HvCMmcwk34C2WNmA"
    },
    {
      "name": "古北1699店",
      "id": "woV3cNDAAA4hCKMY2ml0BaI3iXpQfPdQ"
    },
    {
      "name": "连城广场店",
      "id": "woV3cNDAAAk5q4oy0bgWrexu3lc3yGow"
    },
    {
      "name": "奉发宝龙店",
      "id": "woV3cNDAAAJ5NIYC3-BRbmoJVju7gZQw"
    },
    {
      "name": "中原店",
      "id": "woV3cNDAAAWoMjo1pRXSLm4vcbssAY9Q"
    },
    {
      "name": "丰庄店",
      "id": "woV3cNDAAAkszOK8Bz6b5PnjbWjRVp1Q"
    },
    {
      "name": "西郊乐缤纷店",
      "id": "woV3cNDAAAxbqi1dBKG3EMlV0rjg-g0w"
    },
    {
      "name": "年家浜路店",
      "id": "woV3cNDAAAd3L1kivYW1mBg29oqhtTVA"
    },
    {
      "name": "电台路店",
      "id": "woV3cNDAAAHT1Djf_WwoGmlENJGFCf0w"
    },
    {
      "name": "嘉定日月光店",
      "id": "woV3cNDAAA9IYnW0IwqVHR1xqCPd5vDA"
    },
    {
      "name": "地铁漕河泾店",
      "id": "woV3cNDAAAYn5GVoKWFAA-OJjGvScJOQ"
    },
    {
      "name": "静安大融城店",
      "id": "woV3cNDAAAqxYALko5csy9cromvTZYmQ"
    },
    {
      "name": "灵岩南路店",
      "id": "woV3cNDAAA01KuMeZHEX936Y1XWYzNfQ"
    },
    {
      "name": "翔封路店",
      "id": "woV3cNDAAA56eYBaF_hKR-37kYoGoLBg"
    },
    {
      "name": "七宝宝龙店",
      "id": "woV3cNDAAA9A0mwNwHjFHURHwFhGdKtQ"
    },
    {
      "name": "古方路店",
      "id": "woV3cNDAAAJ5NIYC3-BRbmoJVju7gZQw"
    },
    {
      "name": "共富店",
      "id": "woV3cNDAAAlcuRQ43cNdOlRdhBcrKvnw"
    },
    {
      "name": "亿丰时代广场店",
      "id": "woV3cNDAAA50mV_T6ev8gFaHsJ_zkKRg"
    },
    {
      "name": "松江大学城店",
      "id": "woV3cNDAAAdhyotSCAPwaWt6r3TnrsxA"
    },
    {
      "name": "中环百联店",
      "id": "woV3cNDAAAT-nZVef9lkiQ5Wn1v0KGqw"
    },
    {
      "name": "庆荣店",
      "id": "woV3cNDAAAAY7IT1QHNSsh5NwDylIjyw"
    },
    {
      "name": "东昌路店",
      "id": "woV3cNDAAAQRg4j-RGGKyf31nZSsgnyw"
    },
    {
      "name": "晨阳店",
      "id": "woV3cNDAAALK1s4mqIJ54tMPu8bYfUjw"
    },
    {
      "name": "繁兴店",
      "id": "woV3cNDAAAH8hym5_gfSpu7O77S3tsOg"
    },
    {
      "name": "宝山万达店",
      "id": "woV3cNDAAA0CZPNSsDfmR9UMHk5kzusA"
    },
    {
      "name": "利通广场店",
      "id": "woV3cNDAAArevEKLE9EgwELjyb_10U_Q"
    },
    {
      "name": "青安店",
      "id": "woV3cNDAAAbjDEne_Fb_ApC4uXlESfWA"
    },
    {
      "name": "港城新天地店",
      "id": "woV3cNDAAArGF9o92vWLroPsMjvjWvPQ"
    },
    {
      "name": "古猗园店",
      "id": "woV3cNDAAA37zyJ_-EdKm4bq7FVAnkCA"
    },
    {
      "name": "蓝村店",
      "id": "woV3cNDAAAWbCc9QWDR2XNY0Qbn-a37A"
    },
    {
      "name": "高东园恒路店",
      "id": "woV3cNDAAAix-Um8ipen0LhlCmg-c8FA"
    },
    {
      "name": "裕德店",
      "id": "woV3cNDAAAaMtHslsl93x-DAR2dvGqXg"
    },
    {
      "name": "延长店",
      "id": "woV3cNDAAAiCnG7q_Rwl6xShqiPbo2ow"
    },
    {
      "name": "1点点-唐镇阳光城店",
      "id": "woV3cNDAAA7egxAsS0uTMdMzQXAq4R1g"
    },
    {
      "name": "局门店",
      "id": "woV3cNDAAACO1hYaOej7JbhzE-r-a3SQ"
    },
    {
      "name": "虞姬墩路店",
      "id": "woV3cNDAAAix-Um8ipen0LhlCmg-c8FA"
    },
    {
      "name": "罗宾森店",
      "id": "woV3cNDAAAwS4H9ZOCA_ykD-kfFoK-uw"
    },
    {
      "name": "海事共享区店",
      "id": "woV3cNDAAAPm-Wl-2lr7BSqOtiwbGyYQ"
    },
    {
      "name": "肇嘉浜路店",
      "id": "woV3cNDAAAcdZNXhv5gRBPu_As53gdow"
    },
    {
      "name": "宣中路店",
      "id": "woV3cNDAAAFvIEI8DQOF1SavWJGUOfOA"
    },
    {
      "name": "虹桥大融城",
      "id": "woV3cNDAAAJyIMxJkd1-Cm06oVxg2LWw"
    },
    {
      "name": "青浦奥特莱斯店",
      "id": "woV3cNDAAAYyavMwm5LCX63gZ-4OQG5Q"
    },
    {
      "name": "周浦万达店",
      "id": "woV3cNDAAA_hBIkTSXItgTV9UClH4abA"
    },
    {
      "name": "沈梅东店",
      "id": "woV3cNDAAAz3_iaUDFH23s46Dlzd3-2Q"
    },
    {
      "name": "桃源店",
      "id": "woV3cNDAAA_hqc8ptojcXomz2qxzhwmQ"
    },
    {
      "name": "殷高逸仙店",
      "id": "woV3cNDAAA0_4-OpF-bkIyrb5rg8GMVg"
    },
    {
      "name": "云间新天地店",
      "id": "woV3cNDAAAIrIsQZPeN2-h_eJEiQwBPg"
    },
    {
      "name": "芦潮港店",
      "id": "woV3cNDAAAdBWmbvg-YyC-kP3aW2Stug"
    },
    {
      "name": "图们路店",
      "id": "woV3cNDAAA7egxAsS0uTMdMzQXAq4R1g"
    },
    {
      "name": "川沙百联店",
      "id": "woV3cNDAAAA8Y-JpmhOIZkOdhIBig8Lw"
    },
    {
      "name": "天物空间店",
      "id": "woV3cNDAAActf2_cbpoGazOpB6PSXpFQ"
    },
    {
      "name": "青村振兴店",
      "id": "woV3cNDAAA5yJVrrNppy8VJBFNffjsPQ"
    },
    {
      "name": "国顺东店",
      "id": "woV3cNDAAA7egxAsS0uTMdMzQXAq4R1g"
    },
    {
      "name": "江月店",
      "id": "woV3cNDAAA65EbxePD7-IHCkNh2HGZxw"
    },
    {
      "name": "荣乐西路店",
      "id": "woV3cNDAAAFdgfU3ve-X1cV-eKANYqdw"
    },
    {
      "name": "新金桥太茂店",
      "id": "woV3cNDAAAix-Um8ipen0LhlCmg-c8FA"
    },
    {
      "name": "国和店",
      "id": "woV3cNDAAAsGdpBGhESeMFfkzAs1TEmQ"
    },
    {
      "name": "百色路店",
      "id": "woV3cNDAAAXot5qAe8AZTG9jvH7YaNlQ"
    },
    {
      "name": "松江印象城店",
      "id": "woV3cNDAAAlgFPa9gUt0rSuqfPGf7QCg"
    },
    {
      "name": "新镇路店",
      "id": "woV3cNDAAAFdgfU3ve-X1cV-eKANYqdw"
    },
    {
      "name": "宛平南店",
      "id": "woV3cNDAAAdghcwbnOpjiMOniubxxPzg"
    },
    {
      "name": "同乐店",
      "id": "woV3cNDAAAomtqAm6mLeTt31hBU0QdmQ"
    },
    {
      "name": "悦乐坊店",
      "id": "woV3cNDAAAGES8t0D5K2j7-5qCB2yJpw"
    },
    {
      "name": "通跃店",
      "id": "woV3cNDAAAOwLyj8eIag2ZqsPBK4FucQ"
    },
    {
      "name": "工农南店",
      "id": "woV3cNDAAAKSDOrDaL2HgwPJi-AgLVGw"
    },
    {
      "name": "新村路店",
      "id": "woV3cNDAAAHT1Djf_WwoGmlENJGFCf0w"
    },
    {
      "name": "太阳城店",
      "id": "woV3cNDAAA31yeCortfmJLWXXrDz3hTg"
    },
    {
      "name": "万安店",
      "id": "woV3cNDAAA8iUghbSsKsausvAsISlsXg"
    },
    {
      "name": "河南中路店",
      "id": "woV3cNDAAAu801fhRma5ceUZtAwoQL8Q"
    },
    {
      "name": "青浦吾悦店",
      "id": "woV3cNDAAAfVSFXpqrQjPwDUtpp2PKkg"
    },
    {
      "name": "天等店",
      "id": "woV3cNDAAA7-jpwVeOmwdZC_rw8uFrvA"
    },
    {
      "name": "奉贤南桥店",
      "id": "woV3cNDAAAAbWD6EHieU36A-9XRWeDew"
    },
    {
      "name": "宝乐汇店",
      "id": "woV3cNDAAAMhBXDxxMVsH0kndr_QOJnw"
    },
    {
      "name": "青湖路店",
      "id": "woV3cNDAAA8JvZR1eTpFSRueqvcO9U6g"
    },
    {
      "name": "悠方广场店",
      "id": "woV3cNDAAAuQk0jxe4r-vPxiThKtpOdg"
    },
    {
      "name": "沪亭店",
      "id": "woV3cNDAAAoV2DAp3hNO9HpoQItUmjFg"
    },
    {
      "name": "梅花店",
      "id": "woV3cNDAAA0GRZcFfCwXP2c-KSeTEuLg"
    },
    {
      "name": "聚丰园店",
      "id": "woV3cNDAAAm3FMQzllyjatp6VxD33rdA"
    },
    {
      "name": "仓场路店",
      "id": "woV3cNDAAAqZzCZCrKQfrmXsB0Et46KQ"
    },
    {
      "name": "施湾沃尔玛店",
      "id": "woV3cNDAAACwTEYSO3IbjU6k9kY54_KA"
    },
    {
      "name": "唐镇恒生广场店",
      "id": "woV3cNDAAAix-Um8ipen0LhlCmg-c8FA"
    },
    {
      "name": "塘西店",
      "id": "woV3cNDAAA9UycYANE7Y5qrkp5N7r-6w"
    },
    {
      "name": "长岛店",
      "id": "woV3cNDAAAHT1Djf_WwoGmlENJGFCf0w"
    },
    {
      "name": "闻喜店",
      "id": "woV3cNDAAADqt-6Ttmg-DADvjMfXempQ"
    },
    {
      "name": "崂山店",
      "id": "woV3cNDAAAjHu2w9B27Uw0jxixjrsgig"
    },
    {
      "name": "福海店",
      "id": "woV3cNDAAAKuJeGLJqhhU1Yl-HPYxdrg"
    },
    {
      "name": "奉贤宝龙广场店",
      "id": "woV3cNDAAAAbWD6EHieU36A-9XRWeDew"
    },
    {
      "name": "寺平北店",
      "id": "woV3cNDAAAq_7RI0XbYvIgagJkvWaAAA"
    },
    {
      "name": "新源店",
      "id": "woV3cNDAAAVInMYESqOMvts9tFI_LhDg"
    },
    {
      "name": "横港店",
      "id": "woV3cNDAAAsa1rjCB-s5PbBW1ior-5JA"
    },
    {
      "name": "市北园区店",
      "id": "woV3cNDAAAoQEUrKrc3WEo9O8F7V6cWw"
    },
    {
      "name": "湖滨道店",
      "id": "woV3cNDAAA-1dsKIlxlgq3lhkXZtnnRg"
    },
    {
      "name": "佘山宝乐汇店",
      "id": "woV3cNDAAA8JvZR1eTpFSRueqvcO9U6g"
    },
    {
      "name": "江桥万达店",
      "id": "woV3cNDAAAZ2TpmF96UAGFMSzXVU8wcw"
    },
    {
      "name": "莘朱店",
      "id": "woV3cNDAAAJj2LzvpvNfpEuqm9FG7mpw"
    },
    {
      "name": "华泾天街店",
      "id": "woV3cNDAAA4LPsn1oi13Zp84mi2qSlcg"
    },
    {
      "name": "昌里店",
      "id": "woV3cNDAAASrfB0XPdXyJNtvpPAXhNbw"
    },
    {
      "name": "新溪路店",
      "id": "woV3cNDAAApVKlbijqEif56sCvx34avQ"
    },
    {
      "name": "蒙山路店",
      "id": "woV3cNDAAAARfgEXC4-ItCThF9xAThfg"
    },
    {
      "name": "新颛兴店",
      "id": "woV3cNDAAATZFhCxCJuMv-RTwOqXYC_A"
    },
    {
      "name": "人民东店",
      "id": "woV3cNDAAAqEi4_iYkcS_fY9y_YWAUWQ"
    },
    {
      "name": "凇滨店",
      "id": "woV3cNDAAAoQEUrKrc3WEo9O8F7V6cWw"
    },
    {
      "name": "百联滨江店",
      "id": "woV3cNDAAAboIgj4FE12QbPjQmWjdv7w"
    },
    {
      "name": "中山北路店",
      "id": "woV3cNDAAALGT74HawJeFGYaeXE0dpPQ"
    },
    {
      "name": "方锦店",
      "id": "woV3cNDAAA37zyJ_-EdKm4bq7FVAnkCA"
    },
    {
      "name": "中山广场店",
      "id": "woV3cNDAAAPds1E2QbyuI8JA_bZQ_l3Q"
    },
    {
      "name": "四平店",
      "id": "woV3cNDAAABdddOowftVkB_7QE4CbFew"
    },
    {
      "name": "嘉定宝龙店",
      "id": "woV3cNDAAAix9y9Z7BYmikHM3SMNJ7fw"
    },
    {
      "name": "国华广场店",
      "id": "woV3cNDAAAlgFPa9gUt0rSuqfPGf7QCg"
    },
    {
      "name": "沪松店",
      "id": "woV3cNDAAAvUkhVWajVHNIZ3lsXSERug"
    },
    {
      "name": "顾戴店",
      "id": "woV3cNDAAAdaHssNRTiF76uyJGAO4N0w"
    },
    {
      "name": "闵行欧尚店",
      "id": "woV3cNDAAA7jOIfeQYuzI8fRBD6ReTKw"
    },
    {
      "name": "嘉定百联店",
      "id": "woV3cNDAAAHZ0WcTT9luOH2tjfZw9dlA"
    },
    {
      "name": "金沙江店",
      "id": "woV3cNDAAA_XqhM0EbhMrTNGtgPvGGIw"
    },
    {
      "name": "枫泾中大店",
      "id": "woV3cNDAAAFdgfU3ve-X1cV-eKANYqdw"
    },
    {
      "name": "夏碧店",
      "id": "woV3cNDAAAR0sWwEgJHCZuh451injC1g"
    },
    {
      "name": "城东店",
      "id": "woV3cNDAAAFvIEI8DQOF1SavWJGUOfOA"
    },
    {
      "name": "威海路店",
      "id": "woV3cNDAAAx4ZBWfYkMupBKOxvsI0myQ"
    },
    {
      "name": "赵重公路店",
      "id": "woV3cNDAAAu2QiP6-rweDMS40X7YdSpA"
    },
    {
      "name": "绿港广场店",
      "id": "woV3cNDAAAu2QiP6-rweDMS40X7YdSpA"
    },
    {
      "name": "江浦店",
      "id": "woV3cNDAAAUXW7Phk4ZdobGf4KinO3xw"
    },
    {
      "name": "秀沿店",
      "id": "woV3cNDAAAta5BxQi_az05HkPdErCF0A"
    },
    {
      "name": "古美店",
      "id": "woV3cNDAAADqt-6Ttmg-DADvjMfXempQ"
    },
    {
      "name": "宣化路店",
      "id": "woV3cNDAAAkszOK8Bz6b5PnjbWjRVp1Q"
    },
    {
      "name": "政民店",
      "id": "woV3cNDAAAp0dLnAVx0BKls1ZxY_wwJA"
    },
    {
      "name": "绿苑店",
      "id": "woV3cNDAAAqZzCZCrKQfrmXsB0Et46KQ"
    },
    {
      "name": "LCMart置汇旭辉店",
      "id": "woV3cNDAAAg7ZzyDTjfDemdmpfEqw-sA"
    },
    {
      "name": "环球港店",
      "id": "woV3cNDAAAHT1Djf_WwoGmlENJGFCf0w"
    },
    {
      "name": "嘉定万达店",
      "id": "woV3cNDAAAwfHFcmOj1c2RSVas2jxJhw"
    },
    {
      "name": "尚悦湾店",
      "id": "woV3cNDAAAn1o_36el7BmwAvNNz369ew"
    },
    {
      "name": "明中广场店",
      "id": "woV3cNDAAAoQEUrKrc3WEo9O8F7V6cWw"
    },
    {
      "name": "锦秋店",
      "id": "woV3cNDAAAPM8bVnuf8DdZgDUzayYX-Q"
    },
    {
      "name": "马戏城店",
      "id": "woV3cNDAAAkIShKdRk_P0_N6zV3e-VDQ"
    },
    {
      "name": "金沙和美广场店",
      "id": "woV3cNDAAAlgFPa9gUt0rSuqfPGf7QCg"
    },
    {
      "name": "民耀店",
      "id": "woV3cNDAAA1oHEDuz4KjcyKybzjUVfzA"
    },
    {
      "name": "牡丹江店",
      "id": "woV3cNDAAAGO9sb6p-MCPBGiEf3Aw1cQ"
    },
    {
      "name": "新平高广场店",
      "id": "woV3cNDAAAhydhktJHLcGOz2TW8kc7-w"
    },
    {
      "name": "百熙店",
      "id": "woV3cNDAAAepzHenw5pDUczJoi6RyZaQ"
    },
    {
      "name": "宜山店",
      "id": "woV3cNDAAAokhytdflC8ITgyqxlkUSIQ"
    },
    {
      "name": "陕西南路店",
      "id": "woV3cNDAAAx4ZBWfYkMupBKOxvsI0myQ"
    },
    {
      "name": "水产西路店",
      "id": "woV3cNDAAAta5BxQi_az05HkPdErCF0A"
    },
    {
      "name": "通河路店",
      "id": "woV3cNDAAAMiLOJvgyRvDcl3EG_wb-dw"
    },
    {
      "name": "金山百联店",
      "id": "woV3cNDAAArGF9o92vWLroPsMjvjWvPQ"
    },
    {
      "name": "石门店",
      "id": "woV3cNDAAAKx9_JGx2nHvUFTG5R191EA"
    },
    {
      "name": "飞航广场店",
      "id": "woV3cNDAAA__nXXnRAiWDR0fucb1HX2w"
    },
    {
      "name": "金汇店",
      "id": "woV3cNDAAAKRGReE6s3lxxCwUSscJ8Qw"
    },
    {
      "name": "万泰广场",
      "id": "woV3cNDAAAvp7ClhX5W_MN5KUeYO0Uyw"
    },
    {
      "name": "盛桥店",
      "id": "woV3cNDAAA9Ck8JANMr7PYCpZH9EbWzQ"
    },
    {
      "name": "金山开乐大街店",
      "id": "woV3cNDAAAayDePRoFS0LlgKiYVVUTmw"
    },
    {
      "name": "新徐汇万科店",
      "id": "woV3cNDAAA0I6zO4aZMmBTZMLonjAI_g"
    },
    {
      "name": "宝山正大店",
      "id": "woV3cNDAAA_b4ukH47BAtC948lIjtsqA"
    },
    {
      "name": "临港蓝鲸世界店",
      "id": "woV3cNDAAArGF9o92vWLroPsMjvjWvPQ"
    },
    {
      "name": "三林印象汇店",
      "id": "woV3cNDAAAewJAPRQl51_1YUPpiPQHzA"
    },
    {
      "name": "市台路店",
      "id": "woV3cNDAAAta5BxQi_az05HkPdErCF0A"
    },
    {
      "name": "岚皋路店",
      "id": "woV3cNDAAAZ2JlDtWwUEbkFMF81fpHtg"
    },
    {
      "name": "上海香港广场店",
      "id": "woV3cNDAAA2SiFQA_DslLF_YZmBUTK4Q"
    },
    {
      "name": "协信星光店",
      "id": "woV3cNDAAAOsPBiIYveC8-GpTrZdvumQ"
    },
    {
      "name": "打浦桥日月光店",
      "id": "woV3cNDAAAdk9PCVWcei894niogatvWA"
    },
    {
      "name": "北园路店",
      "id": "woV3cNDAAA37FfI5J3eZj9W1xyhe3Psg"
    },
    {
      "name": "复地活力城店",
      "id": "woV3cNDAAAdZBhrm4JGdH5-8QVNv_ICQ"
    },
    {
      "name": "临港电机学院店",
      "id": "woV3cNDAAAHg0Afj_7SEBXDcNlosyG0w"
    },
    {
      "name": "1点点友邦金融中心店",
      "id": "woV3cNDAAAeYkRWMeN8GM0xtgzTFXdqg"
    },
    {
      "name": "蔡伦路店",
      "id": "woV3cNDAAArLbPeamcvGRPsrC-Mt-hpw"
    },
    {
      "name": "1点点光启城店",
      "id": "woV3cNDAAAhunmOzfJLtyfQl6pJuM92w"
    },
    {
      "name": "汇锦里店",
      "id": "woV3cNDAAAu6e0GEwk3loZExZ4sU2T9w"
    },
    {
      "name": "中科路店",
      "id": "woV3cNDAAAgYQC4ptoKG-UnTZrdA07Yg"
    },
    {
      "name": "北艾路店",
      "id": "woV3cNDAAARFak59DX57vQMijTEhp7wg"
    },
    {
      "name": "四平电影院店",
      "id": "woV3cNDAAAeYx3tpC7gg7FyCNZKiFeYw"
    },
    {
      "name": "肖塘店",
      "id": "woV3cNDAAAEyBg09NvNAsp3aaPkvwjgg"
    },
    {
      "name": "川沙商业广场店",
      "id": "woV3cNDAAA7egxAsS0uTMdMzQXAq4R1g"
    },
    {
      "name": "城大路店",
      "id": "woV3cNDAAAcqh90mLSpZriBwnF1h6OLg"
    },
    {
      "name": "龙湖虹桥天街店",
      "id": "woV3cNDAAAZhR0L_rHaRAIVuY_mLt_-g"
    },
    {
      "name": "张江大橘天地店",
      "id": "woV3cNDAAAdKf_irF6hoFpea71sIt7hQ"
    },
    {
      "name": "云翔未来城店",
      "id": "woV3cNDAAAeAESNeOn6jm0xKv5YrRB4g"
    },
    {
      "name": "松江新城地铁店",
      "id": "woV3cNDAAAYpZfOnUxkaLdrqm4XIAnQQ"
    },
    {
      "name": "新飞洲国际店",
      "id": "woV3cNDAAA5goGZRmozAlvFDvbmuOZVg"
    },
    {
      "name": "吴江路店",
      "id": "woV3cNDAAAwixqtqzG7yo1yE97UeltLA"
    },
    {
      "name": "绿地乐和城店",
      "id": "woV3cNDAAA72fLs5jGq34Jg0Ja_gUp1Q"
    },
    {
      "name": "法华镇定西店",
      "id": "woV3cNDAAAJyIMxJkd1-Cm06oVxg2LWw"
    },
    {
      "name": "浦建巴黎春天店",
      "id": "woV3cNDAAAdyI2-Chgg_9L67jVyV2sWQ"
    },
    {
      "name": "奉贤海旗店",
      "id": "woV3cNDAAAPm-Wl-2lr7BSqOtiwbGyYQ"
    },
    {
      "name": "纬地店",
      "id": "woV3cNDAAAkK4S9rSjeHcIOcAdOp2Rrw"
    },
    {
      "name": "中山北一店",
      "id": "woV3cNDAAAMV9wUvHcF_kaPO6wSzypfA"
    },
    {
      "name": "沪亭南店",
      "id": "woV3cNDAAAvUkhVWajVHNIZ3lsXSERug"
    },
    {
      "name": "新庆荣店",
      "id": "woV3cNDAAAu_dV_CDErrQmxslemCYSFA"
    },
    {
      "name": "松江大橘店",
      "id": "woV3cNDAAAIZGTgjHvvhHTuymmhDdwWQ"
    },
    {
      "name": "上理工思餐厅三楼店",
      "id": "woV3cNDAAARA0Icy9zHMpAvCVvC8xeQw"
    },
    {
      "name": "亭林名悦广场店",
      "id": "woV3cNDAAAi6oNVV812G0ftQA1ZJxd-g"
    },
    {
      "name": "新江川北",
      "id": "woV3cNDAAAdhyotSCAPwaWt6r3TnrsxA"
    },
    {
      "name": "趣开专属店",
      "id": "woV3cNDAAAYIPIEAQqcYgnI6HA2ZJJ_g"
    },
    {
      "name": "人民东店1",
      "id": "woV3cNDAAAWJlLadU9iNiugGGlFbc8qg"
    },
    {
      "name": "1点点凯德虹口商业中心店1",
      "id": "woV3cNDAAAUOWSYW7Xb_mUvnb6QNg4FA"
    },
    {
      "name": "盈港路富绅时代店（老）",
      "id": "woV3cNDAAAOGzDtmPW3EZPtFPVrBwl2g"
    },
    {
      "name": "绿都绣云里店",
      "id": "woV3cNDAAAlTAiYV6MbIfRp46ROikP9g"
    },
    {
      "name": "奉贤上师大金桂苑店",
      "id": "woV3cNDAAADA33_Sl-qr-1JFPc1LX2XQ"
    },
    {
      "name": "中集美兰湖金地广场店",
      "id": "woV3cNDAAAfPSfiE4r3CPSFYsfdf1E9g"
    },
    {
      "name": "CP静安店",
      "id": "woV3cNDAAA6D-U97WCMZ7V9-MCpQ8yMA"
    },
    {
      "name": "二工大西一食堂店",
      "id": "woV3cNDAAA24MHYq47T5dXiPnEYfjy3w"
    },
    {
      "name": "贤达国交学院楼店",
      "id": "woV3cNDAAA5-9H9g7_SNPQOg0rq78m-Q"
    },
    {
      "name": "江桥万达金街5号门店",
      "id": "woV3cNDAAAZ2TpmF96UAGFMSzXVU8wcw"
    },
    {
      "name": "唐镇印象汇店",
      "id": "woV3cNDAAA4Qe-TzWuW_lRCY3tXI0OQQ"
    },
    {
      "name": "奉贤上应大一食堂二楼店",
      "id": "woV3cNDAAAXA9fzmozT5_scTqfktR0tw"
    },
    {
      "name": "龙华会店",
      "id": "woV3cNDAAAVInMYESqOMvts9tFI_LhDg"
    },
    {
      "name": "陆悦天地店",
      "id": "woV3cNDAAAoP82ezncAKdsLiRPNzVnEA"
    },
    {
      "name": "新浦三路店",
      "id": "woV3cNDAAAINvkPeo1bCBD1Qa-1McUFw"
    },
    {
      "name": "听悦店",
      "id": "woV3cNDAAA74OtXSWT8E7xNfn5uZxCYw"
    },
    {
      "name": "长宁来福士店",
      "id": "woV3cNDAAA2dxtVPfVydm6nTdmVER3jQ"
    },
    {
      "name": "大林店",
      "id": "woV3cNDAAAjgoou8Klm2K5-7naJ1byDA"
    },
    {
      "name": "前湾印象城店",
      "id": "woV3cNDAAAgE5H7KJDQouF8GgWNtt7SQ"
    },
    {
      "name": "展讯中心店",
      "id": "woV3cNDAAAY8J4ZJGXueDNkou2icTVyg"
    },
    {
      "name": "嘉园坊店",
      "id": "woV3cNDAAAEyBg09NvNAsp3aaPkvwjgg"
    },
    {
      "name": "前滩企业天地店",
      "id": "woV3cNDAAA41vSYr8WEj6SYu9plaORJA"
    },
    {
      "name": "青浦新桥路店",
      "id": "woV3cNDAAA3zQUiCfCtQjmDWiJnslzew"
    },
    {
      "name": "东华大学店",
      "id": "woV3cNDAAAsIXXZZBONqB6v0Q1rZvJcQ"
    },
    {
      "name": "南汇大学城店",
      "id": "woV3cNDAAA2phdDfmDQE4QUk52UoTSCQ"
    },
    {
      "name": "奉贤瓦洪店",
      "id": "woV3cNDAAAgqKJBSoN5dlfO2PoNO6ltg"
    },
    {
      "name": "安中路店",
      "id": "woV3cNDAAAayDePRoFS0LlgKiYVVUTmw"
    }
  ]
};

// 存储所有汇总数据
let summaryData = [];

/**
 * 创建输出目录
 */
function createOutputDir() {
  if (!fs.existsSync(config.outputDir)) {
    fs.mkdirSync(config.outputDir, { recursive: true });
    console.log(`📁 创建输出目录: ${config.outputDir}`);
  }
}

/**
 * 并发池
 * limit: 最大并发数
 * items: 要处理的数据列表
 * iteratorFn: 处理函数
 */
async function asyncPool(limit, items, iteratorFn) {
  const results = [];
  const executing = [];

  for (const item of items) {
    const p = Promise.resolve().then(() => iteratorFn(item));
    results.push(p);

    const e = p.finally(() => {
      const index = executing.indexOf(e);
      if (index >= 0) executing.splice(index, 1);
    });

    executing.push(e);

    if (executing.length >= limit) {
      await Promise.race(executing);
    }
  }

  return Promise.all(results);
}

/**
 * 调用接口获取单页数据
 * @param {number} page 页码
 * @param {string} followerUserId 跟进人ID
 * @returns {Promise<{data: Array, totalPage: number}>}
 */
async function fetchPageData(page, followerUserId) {
  try {
    const response = await axios.post(
      config.baseUrl,
      {
        page: page,
        computeType: 1,
        pageSize: config.pageSize,
        followerUserId: followerUserId,
        tagIdList: [],
        addTime: config.addTime
      },
      {
        headers: config.headers,
        decompress: true,
        timeout: 20000
      }
    );

    if (response.data.success && response.data.data) {
      const { list = [], totalPage = 0 } = response.data.data;

      const pageData = list.map(item => ({
        name: item.name || '',
        externalUserId: item.externalUserId || '',
        innerUnionId: item.innerUnionId || '',
        staffUserId: item.staffUserId || '',
        shopId: item.shopId || '',
        state: item.state || ''
      }));

      return {
        data: pageData,
        totalPage
      };
    } else {
      console.error(`❌ 获取数据失败: ${response.data.message || '未知错误'}，页码: ${page}`);
      return {
        data: [],
        totalPage: 0
      };
    }
  } catch (error) {
    console.error(`❌ 请求出错: ${error.message}，页码: ${page}`);
    return {
      data: [],
      totalPage: 0
    };
  }
}

/**
 * 将数据写入Excel文件
 * @param {Array} data 要写入的数据
 * @param {string} fileName 文件名
 * @param {string} sheetName 工作表名称
 */
function writeToExcel(data, fileName, sheetName = '客户数据') {
  try {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    XLSX.writeFile(workbook, fileName);
    console.log(`📊 数据已写入文件: ${fileName}`);
  } catch (error) {
    console.error(`❌ 写入Excel失败: ${error.message}`);
  }
}

/**
 * 获取指定门店的所有数据（分页并发）
 * @param {Object} followerUser 包含name和id的对象
 * @returns {Promise<Array>}
 */
async function fetchAllDataForFollower(followerUser) {
  console.log(`\n========================================`);
  console.log(`🔍 开始获取【${followerUser.name}】的数据 (ID: ${followerUser.id})`);
  console.log(`========================================`);

  // 先查第一页，拿到总页数
  const firstPageResult = await fetchPageData(1, followerUser.id);
  const totalPage = firstPageResult.totalPage || 0;

  let allData = [...firstPageResult.data];

  if (totalPage <= 1) {
    console.log(`📄 【${followerUser.name}】只有1页数据，共 ${allData.length} 条`);
    return allData;
  }

  console.log(`📄 【${followerUser.name}】总共有 ${totalPage} 页数据`);

  // 生成剩余页码
  const restPages = Array.from({ length: totalPage - 1 }, (_, i) => i + 2);

  // 分页并发请求
  const pageResults = await asyncPool(config.pageConcurrency, restPages, async (page) => {
    console.log(`📥 正在获取【${followerUser.name}】第 ${page}/${totalPage} 页...`);
    return await fetchPageData(page, followerUser.id);
  });

  for (const result of pageResults) {
    allData.push(...result.data);
  }

  console.log(`✅ 【${followerUser.name}】数据获取完成，共 ${allData.length} 条`);
  return allData;
}

/**
 * 处理单个门店：拉数据 + 导出单独Excel + 返回汇总数据
 */
async function processFollower(followerUser) {
  const userData = await fetchAllDataForFollower(followerUser);

  if (userData.length === 0) {
    console.log(`⚠️ 【${followerUser.name}】没有获取到任何数据`);
    return [];
  }

  // 保留你原来的“每个门店导出一个Excel”
  const fileName = path.join(config.outputDir, `${followerUser.name}_客户数据.xlsx`);
  writeToExcel(userData, fileName, `${followerUser.name}数据`);

  // 返回用于汇总的数据
  return userData.map(item => ({
    门店名称: followerUser.name,
    ...item
  }));
}

/**
 * 主函数 - 执行完整的导出流程（门店并发）
 */
async function main() {
  console.log('🚀 开始执行客户数据导出任务');
  console.log(`📋 共需要处理 ${config.followerUserList.length} 个跟进人`);
  console.log(`⚙️ 门店并发: ${config.storeConcurrency}，分页并发: ${config.pageConcurrency}`);

  createOutputDir();

  // 门店并发执行
  const allResults = await asyncPool(
    config.storeConcurrency,
    config.followerUserList,
    async (followerUser) => {
      return await processFollower(followerUser);
    }
  );

  // 汇总所有门店数据
  summaryData = allResults.flat();

  console.log(`\n========================================`);
  console.log(`📝 开始生成汇总文件`);
  console.log(`========================================`);

  // 保留你原来的“总汇总导出”
  const summaryFilePath = path.join(config.outputDir, config.summaryFileName);
  writeToExcel(summaryData, summaryFilePath, '所有门店汇总');

  console.log(`\n🎉 所有任务执行完成！`);
  console.log(`📊 汇总数据共 ${summaryData.length} 条`);
  console.log(`📁 所有文件已保存到: ${config.outputDir}`);
}

// 执行主函数
main().catch(error => {
  console.error(`💥 程序执行出错: ${error.message}`);
  process.exit(1);
});