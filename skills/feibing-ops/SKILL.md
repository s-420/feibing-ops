---
name: feibing-ops
description: 处理飞冰业务需求时使用——识别客户需求、查 taxonomy 定位、复用或新写脚本、dryRun、按 8 项约定自动 commit/push、记台账与留痕。触发词：飞冰、处理客户需求、批量替换、群链接、欢迎语、桌台、门店、渠道码、建店、导出、激活、退款。
---

# 飞冰业务操作自动化（feibing-ops）

## 用途

把对接客户问题的繁琐操作，按飞冰代码库的 Center 架构封装成规范化脚本 / skill，并自动提交到共享仓库。目标：遇到客户需求，agent 沿固定路径自动定位、执行、封装、提交、留痕，不再靠人现场判断。

## 前置

- 仓库：`D:\Code\feibing\feibing-ops`（已 git 化，remote = git@github.com:s-420/feibing-ops.git）
- 规范：`CONTRIBUTING.md`（8 项约定）
- 分类清单：`taxonomy/index.md`（唯一事实源：Center 架构 + 对象→Center 映射 + 脚本登记）

## 工作流程

遇到飞冰业务需求时，按以下步骤走：

### 1. 登记需求

用 `templates/需求登记.md` 明确四要素：**客户 / 对象 / 动作 / 期望结果**。

### 2. 查 taxonomy 定位

读 `taxonomy/index.md`：
- 识别**对象**（桌台/群/门店/客户/渠道码/企微…）→ 查「对象→Center 映射」得到 Center；
- 识别**动作**（导出/批量写/匹配/排查/建实体）→ 得到风险等级。

### 3. 分支决策

- **命中已有脚本** → 复用执行；
- **未命中** → 按封装分层新增：
  - 输入输出确定 → 写**脚本**，放 `scripts/<center>/`，命名 `对象-动作-说明.js`，带元数据头；
  - 需判断、多步组合 → 写**skill**，放 `skills/<center>/`；
  - 只能界面手点 → 写**SOP 文档**，放 `skills/<center>/`。

### 4. dryRun 与脱敏

- 写操作脚本必须先 **dryRun 预演**（改线上数据前备份 + 小范围试跑）；
- 检查**无硬编码 token**：`eyJ...` 令牌段必须替换为 `process.env.VINCI_AUTHORIZATION`（或占位符）；
- `.env`、`input/`、`output/` 等业务数据与凭据**不提交**。

### 5. 自动提交

- `git add -A` → commit，message 格式 `[Center] 动作：说明`（如 `[sc] 新增：桌台批量替换欢迎语脚本`）；
- 本机直接提交 main 并 `git push`（小组其他成员协作时走 feature 分支 + PR）；
- 同步更新 `taxonomy/index.md`（登记新脚本 / 新条目）。

### 6. 留痕（双层）

- 仓库 `records/YYYY-MM-DD/<客户>-<对象>-<动作>.md` 记一条业务留痕（用 `templates/留痕记录.md`）；
- 本机台账 `D:\AgentRules\ledger\YYYY-MM-DD.md` 追加一条（6 字段：时间 / 目标 / 使用 agent / 状态 / 产物 / 待办遗漏）。

## 安全约束（硬性，违反即不合入）

1. 绝不提交 `.env`、token、客户业务数据（门店/群/账号/手机号）；
2. 写操作脚本无 dryRun 不合入；
3. 客户不进目录——能力与数据分离，客户只是运行参数。

## 规范速查

- 目录：`scripts/<center>/`、`skills/<center>/`（center ∈ sc / ac / mc / tc / xc）
- 命名：`对象-动作-说明`
- 元数据头：`@对象 / @动作 / @风险 / @输入 / @输出 / @验收`
- 风险：导出、匹配 = 低（直接跑）；批量写 = 高（dryRun + 备份 + 复核 + 回滚）；建实体 = 中（先 SOP 确认）
- commit：`[Center] 动作：说明`

## Center 速查（飞冰后端 10 个微服务）

| Center | 职责 |
|---|---|
| sc 商家中心 | 门店、桌台、会员、标签、欢迎语、渠道码、优惠券 |
| ac 活动获客 | 活动、加好友、客户运营、群欢迎语、活码、营销话术 |
| mc 营销中心 | 优惠券、团购、打卡、抽奖、积分、任务 |
| tc 交易中心 | 订单、退款、支付、履约 |
| xc 微信中心 | 企微、公众号、小程序、开放平台 |
