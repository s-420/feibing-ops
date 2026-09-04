# feibing-ops · 飞冰业务操作库

> 飞冰（Feibing）业务操作封装库——把日常对接客户的繁琐操作，封装成规范化的脚本 / skill 资产库。
> 是通用模板 [`ops-template`](https://github.com/s-420/ops-template) 的第一个落地实例。

## 飞冰是什么

以企业微信为底座、面向线下连锁品牌（核心餐饮）的「新零售业务操作系统」：
飞冰引擎 SCRM（客户运营）+ 飞冰互联（渠道协同）+ 飞冰接龙（社群营销）。API 见 `https://vinci-api.feibing.tech/sc/swagger-ui.html`。

## 本库解决什么

对接客户的需求表面杂乱（换群链接、改欢迎语、建店、查权限、激活账号……），
本库用「对象 × 动作」分类把它们收敛成可复用、可交接的资产，让人和 AI 都能稳定检索与执行。

## 飞冰业务对象（7 个）

| 对象 | 说明 |
|---|---|
| 门店 | 建店、门店配置、门店号、战区、客服号、入群链接 |
| 桌台 | 桌台码、桌台客服、欢迎语、桌台标签、桌台拷贝 |
| 群 | 客户群、裂变群、LBS 群、入群链接 |
| 客户 | 客户列表、标签、画像、优惠券、退款 |
| 企微 | 成员、账号激活、推送、授权、应用续费 |
| 渠道码 | 渠道活码、LBS 码、店外绑定码 |
| 活动 | 接龙、朋友圈任务、推送 |

> 客户（茶瀑布 / 沪上阿姨 / 一点點 / 有熊的花园 / 谷谷雪山 …）是**数据参数**，不进目录。

## 目录结构

```
feibing-ops/
├── README.md            # 本文件
├── CONTRIBUTING.md      # 上传规范（8 项约定，继承自 ops-template）
├── taxonomy/index.md    # 分类清单：飞冰对象清单 + agent 路由（唯一事实源）
├── skills/              # skill 库（按对象分目录）
├── scripts/             # 脚本库（按对象分目录）
├── records/             # 业务处理留痕
└── templates/           # 模板（继承自 ops-template）
```

## 给 agent 的入口（接手即读）

1. `taxonomy/index.md` —— 先建立「飞冰有哪些业务对象、怎么路由」的认知；
2. `CONTRIBUTING.md` —— 掌握上传与封装的 8 项约定；
3. 对应对象的 `skills/`、`scripts/` —— 复用已有能力。

## 脚本归位状态

已有脚本位于 `D:\Code\feibing\store-data-extractor\store-data-extractor\src\`（55 个，拼音命名），
**尚未归位到本库**。归位计划见 `taxonomy/index.md` 第六节。
归位时：复制 → 按 `对象-动作-说明` 重命名 → 凭据走 `.env`（不提交）→ 写操作补 dryRun。

## 凭据说明（重要）

脚本调用 vinci-api 的登录令牌有效期约 24 小时，存于 `.env` 的 `VINCI_AUTHORIZATION`，**永不提交**。
