# feibing-ops · 飞冰业务操作库

> 飞冰（Feibing）业务操作封装库——把日常对接客户的繁琐操作，封装成规范化的脚本 / skill 资产库。
> 是通用模板 [`ops-template`](https://github.com/s-420/ops-template) 的第一个落地实例。

## 飞冰是什么

以企业微信为底座、面向线下连锁品牌（核心餐饮）的「新零售业务操作系统」：
飞冰引擎 SCRM（客户运营）+ 飞冰互联（渠道协同）+ 飞冰接龙（社群营销）。API 见 `https://vinci-api.feibing.tech/sc/swagger-ui.html`。

## 本库解决什么

对接客户的需求表面杂乱（换群链接、改欢迎语、建店、查权限、激活账号……），
本库按飞冰代码库的真实功能域（Center）分类，把它们收敛成可复用、可交接的资产，让人和 AI 都能稳定检索与执行。

## 飞冰功能域（10 个 Center，对齐后端代码库）

| Center | 全称 | 职责 |
|---|---|---|
| `sc` | Seller Center 商家中心 | 门店、桌台、员工、会员、标签、欢迎语、渠道码/二维码、优惠券 |
| `ac` | AC Center 活动/获客 | 活动、加好友、客户运营、群欢迎语、活码、营销话术 |
| `mc` | Market Center 营销中心 | 优惠券、团购、打卡、抽奖、积分、任务（飞冰接龙） |
| `tc` | Trade Center 交易中心 | 订单、退款、支付、履约 |
| `xc` | Weixin Center 微信中心 | 企微、公众号、小程序、开放平台 |
| `uc`/`bc`/`cc`/`pc`/`dc` | 用户/基础/内容/平台/数据 | 认证、消息、内容、权限、飞书 |

> 客户（茶瀑布 / 沪上阿姨 / 一点點 / 有熊的花园 / 谷谷雪山 …）是**数据参数**，不进目录。
> 业务对象 → Center 的映射见 `taxonomy/index.md` 第四节。

## 目录结构

```
feibing-ops/
├── README.md            # 本文件
├── CONTRIBUTING.md      # 上传规范（8 项约定，继承自 ops-template）
├── taxonomy/index.md    # 分类清单：飞冰 Center 清单 + agent 路由（唯一事实源）
├── skills/              # skill 库（按 Center 分目录）
├── scripts/             # 脚本库（按 Center 分目录）
├── records/             # 业务处理留痕
└── templates/           # 模板（继承自 ops-template）
```

## 给 agent 的入口（接手即读）

1. `taxonomy/index.md` —— 先建立「飞冰有哪些功能域（Center）、怎么路由」的认知；
2. `CONTRIBUTING.md` —— 掌握上传与封装的 8 项约定；
3. 对应 Center 的 `skills/`、`scripts/` —— 复用已有能力。

## 脚本归位状态

已有脚本位于 `D:\Code\feibing\store-data-extractor\store-data-extractor\src\`（55 个，拼音命名），
**尚未归位到本库**。归位计划见 `taxonomy/index.md` 第八节。
归位时：复制到 `scripts/<center>/` → 按 `对象-动作-说明` 重命名 → 凭据走 `.env`（不提交）→ 写操作补 dryRun。

## 凭据说明（重要）

脚本调用 vinci-api 的登录令牌有效期约 24 小时，存于 `.env` 的 `VINCI_AUTHORIZATION`，**永不提交**。
