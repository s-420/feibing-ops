# 分类清单（Taxonomy）— 飞冰实例

> **唯一事实源**。本库所有脚本 / skill 按飞冰代码库的真实功能域（Center）分类。
> agent 接手飞冰业务，**先读本文件**建立业务认知，再按「路由引导」定位能力。
> 通用规范见 `ops-template`：<https://github.com/s-420/ops-template>

## 一、分类方法（Center × 对象）

飞冰后端是 10 个 Center 的微服务架构（源码 `D:\Code\feibing\feibing-project\backend\`）。本库一级分类对齐 Center：

- **Center**（功能域）：飞冰微服务，决定 `skills/`、`scripts/` 的**一级目录**，回答「归属飞冰哪个子系统、调哪个服务的接口」。
- **对象**（业务对象）：门店 / 桌台 / 群 / 客户……，做**二级分类 / 元数据字段**，回答「操作哪个业务对象」。
- **客户**（茶瀑布 / 沪上阿姨 / 一点點 / 有熊的花园 / 谷谷雪山 …）：输入参数，**不进目录**。

## 二、agent 路由引导（接手即用）

拿到客户需求，按 4 步走：

1. **识别对象**：说的是哪个业务对象？（门店？桌台？群？客户？渠道码？）
2. **查映射**：在「对象 → Center 映射」里找到它属于哪个 Center。
3. **识别动作**：导出 / 批量写 / 匹配 / 排查？
4. **定位能力**：到对应 Center 目录找脚本 / skill；未命中 → 按 `CONTRIBUTING.md` 第 5 条封装。

> 判断优先级：**先对象 → 映射 Center → 再动作 → 最后客户**。

## 三、飞冰 Center 清单（源码真实划分）

| Center | 全称 | 职责（controller 归纳） | 脚本落点 |
|---|---|---|---|
| `sc` | Seller Center 商家中心 | 门店、桌台、员工、会员、标签、欢迎语、渠道码/二维码、优惠券、订单、品牌、权限 | ★核心 |
| `ac` | AC Center 活动/获客 | 活动、加好友、客户运营、朋友圈、群欢迎语、活码、营销话术、自动标签 | ★核心 |
| `mc` | Market Center 营销中心 | 优惠券、团购、打卡、抽奖、积分、任务、秒杀、盲盒（飞冰接龙） | ★ |
| `tc` | Trade Center 交易中心 | 购物车、订单、退款、支付、履约 | ★ |
| `xc` | Weixin Center 微信中心 | 企微（Cp*）、公众号、小程序、开放平台、第三方平台 | ★ |
| `uc` | User Center 用户中心 | 用户、账户、认证、团队 | 基础 |
| `bc` | Base Center 基础中心 | 消息、存储、短链、标签、模板、打印 | 基础 |
| `cc` | Content Center 内容中心 | 内容、足迹、评论、空间 | 基础 |
| `pc` | Platform Center 平台中心 | 权限、分类、配置 | 基础 |
| `dc` | DC Center | 飞书对接 | 基础 |

## 四、对象 → Center 映射（业务对象归位到功能域）

| 业务对象 | 归属 Center |
|---|---|
| 门店 | `sc` |
| 桌台 | `sc` |
| 群（客户群 / 裂变群 / LBS 群） | `sc` + `ac` + `xc` |
| 客户 / 会员 | `ac` + `sc` |
| 企微（成员 / 授权 / 推送） | `xc` |
| 渠道码 / 二维码 | `sc` |
| 活动 / 接龙 | `mc` + `ac` |
| 退款 / 订单 | `tc` |

## 五、脚本登记（按 Center，来自 store-data-extractor，待归位）

| Center | 已有脚本（store-data-extractor，拼音原名） |
|---|---|
| `sc-商家中心` | DaoChuQuanBuMenDian、DaoChuZhiDinMenDianZuoTai、PiLiangTongBuMenDianPeiZhi、TiHuanMenDianKeFuHao、TiHuanMenDianRuQunLianJie、XiuGaiWeiXinMenDianHaoLink、PiLiangGengHuanZuoTaiLBSQunLianJie、PiLiangTiHuanZuoTaiHuanYingYuNeiRong、PiLiangXiuGaiZuoTaiDuiYinKeFu、PiLiangXiuGaiZuoTaiZiDingYiBiaoQian、PiLiangKaoBeiZuoTai、XunHuanChuangJianZuoTai、DaoChuQUanBuKeHuQun、DaoChuQuanBuLieBianQun、PiPeiLieBianQunHeKeHuQunShaiXuan、PiPeiLieBianQunHeQuanMenDian、DaoChuQuanBuLBSMa、DaoChuDianWaiBangDingMa、PiLiangChuangJianQuDaoMa、deleteWuYongQuDao、ShanChuWuYongQuDaoMaHuLian、DaoChuQuanBuYouHuiQuan |
| `xc-微信中心` | DaoChuQuanBuChengYuan、PiLiangJiHuoZhangHu、PiLiangShuaXinMenDianKeFuYiDuiYi |
| `tc-交易中心` | PiLiangTuiKuan |
| `ac-活动获客` | （暂无脚本，多为界面操作：朋友圈任务、加好友、群欢迎语） |

## 六、动作与风险元数据（固定表）

| 动作 | 说明 | 风险 | 执行要求 |
|---|---|---|---|
| 导出 / 查询 | 只读，拉数据 / 生成报表 | 低 | 直接跑 |
| 匹配 / 比对 | 计算、筛选、对账 | 低 | 直接跑，结果留痕 |
| 批量写 | 批量增 / 改 / 删 | 高 | dryRun → 备份 → 试跑 → 执行 → 复核 |
| 排查 / 诊断 | 只查不动，定位问题 | 低 | 走决策树，产出结论 |
| 建实体 | 建店 / 建群 / 建码 | 中 | 先 SOP 确认，再执行 |

## 七、本清单维护规则

- 新增能力 → 先定位到 Center，再在对应 Center 目录 + 本节「脚本登记」里登记。
- 脚本归位 → 从 `store-data-extractor` 复制到 `scripts/<center>/`，按 `对象-动作-说明` 重命名，凭据走 `.env`，写操作补 dryRun。
- 本文件是唯一事实源，脚本 / skill 与清单不一致时，以清单为准并修正脚本。

## 八、脚本归位状态

- [ ] sc-商家中心 脚本归位
- [ ] xc-微信中心 脚本归位
- [ ] tc-交易中心 脚本归位
- [ ] ac-活动获客 脚本归位（当前无脚本，先建 SOP 文档）
