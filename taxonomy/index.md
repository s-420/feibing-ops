# 分类清单（Taxonomy）— 飞冰实例

> **唯一事实源**。本库所有脚本 / skill 都在这里登记，按「对象 × 动作」分类。
> agent 接手飞冰业务，**先读本文件**建立业务认知，再按「路由引导」定位能力。
> 通用规范见 `ops-template`：<https://github.com/s-420/ops-template>

## 一、分类方法（对象 × 动作）

- **对象**：飞冰业务对象（门店 / 桌台 / 群 / 客户 / 企微 / 渠道码 / 活动）→ 决定 `skills/`、`scripts/` 的**一级目录**。
- **动作**：导出 / 批量写 / 匹配 / 排查 → 决定**元数据字段**与风险等级。
- **客户**（茶瀑布 / 沪上阿姨 / 一点點 / 有熊的花园 / 谷谷雪山 / 轻享 / 湖南领慧 …）：输入参数，**不进目录**。

## 二、agent 路由引导（接手即用）

拿到一个客户需求（一句话 / 一段话），按以下 4 步走：

1. **识别对象**：这句话说的是哪个业务对象？（门店？桌台？群？客户？企微？渠道码？活动？）
2. **识别动作**：要对这个对象做什么？（导出 / 批量写 / 匹配 / 排查？）
3. **查清单**：在下方「对象清单」里找 `对象 × 动作` 对应条目。
4. **定位能力**：
   - 命中 → 复用条目里登记的脚本 / skill；
   - 未命中 → 按 `CONTRIBUTING.md` 第 5 条选封装层级，新增条目。

> 判断优先级：**先对象（放哪）→ 再动作（怎么跑）→ 最后客户（给谁跑）**。

## 三、动作与风险元数据（固定表）

| 动作 | 说明 | 风险 | 执行要求 |
|---|---|---|---|
| 导出 / 查询 | 只读，拉数据 / 生成报表 | 低 | 直接跑 |
| 匹配 / 比对 | 计算、筛选、对账 | 低 | 直接跑，结果留痕 |
| 批量写 | 批量增 / 改 / 删 | 高 | dryRun → 备份 → 试跑 → 执行 → 复核 |
| 排查 / 诊断 | 只查不动，定位问题 | 低 | 走决策树，产出结论 |
| 建实体 | 建店 / 建群 / 建码 | 中 | 先 SOP 确认，再执行 |

## 四、对象清单（飞冰）

> 「已有能力」列引用 `store-data-extractor` 里的脚本（拼音原名，待归位到本库 `scripts/<对象>/`）；「待封装」列是已识别但未脚本化的操作。

| 对象 | 说明 | 常见动作 | 已有能力（store-data-extractor） | 待封装 |
|---|---|---|---|---|
| 门店 | 建店、门店配置、门店号、战区、客服号、入群链接 | 建店、导出、批量写 | DaoChuQuanBuMenDian、PiLiangTongBuMenDianPeiZhi、TiHuanMenDianKeFuHao、TiHuanMenDianRuQunLianJie、XiuGaiWeiXinMenDianHaoLink | 新建店铺 SOP、账号激活 |
| 桌台 | 桌台码、桌台客服、欢迎语、桌台标签、桌台拷贝 | 导出、批量写 | DaoChuZhiDinMenDianZuoTai、PiLiangGengHuanZuoTaiLBSQunLianJie、PiLiangTiHuanZuoTaiHuanYingYuNeiRong、PiLiangXiuGaiZuoTaiDuiYinKeFu、PiLiangXiuGaiZuoTaiZiDingYiBiaoQian、PiLiangKaoBeiZuoTai、XunHuanChuangJianZuoTai | 桌台欢迎语恢复 |
| 群 | 客户群、裂变群、LBS 群、入群链接 | 导出、匹配、排查、建实体 | DaoChuQUanBuKeHuQun、DaoChuQuanBuLieBianQun、PiPeiLieBianQunHeKeHuQunShaiXuan、PiPeiLieBianQunHeQuanMenDian | 群重复排查（业务重名 vs 数据重复）、创建裂变群码 |
| 客户 | 客户列表、标签、画像、优惠券、退款 | 导出、批量写 | DaoChuKeHuLieBiao、DaoChuKeHuEidXinXi、DaoChuKeHuEidToUidXinXi、DaoChuQuanBuYouHuiQuan、PiLiangTuiKuan | 会员身份识别排查 |
| 企微 | 成员、账号激活、推送、授权、应用续费 | 导出、批量写、排查 | DaoChuQuanBuChengYuan、PiLiangJiHuoZhangHu、PiLiangShuaXinMenDianKeFuYiDuiYi | 81013 报错排查、组织架构权限、批量关联群权限 |
| 渠道码 | 渠道活码、LBS 码、店外绑定码 | 导出、批量写 | DaoChuQuanBuLBSMa、DaoChuDianWaiBangDingMa、PiLiangChuangJianQuDaoMa、deleteWuYongQuDao、ShanChuWuYongQuDaoMaHuLian | — |
| 活动 | 接龙、朋友圈任务、推送 | 建实体、批量写 | （无脚本，多为界面操作） | 推送朋友圈任务、接龙活动 |

## 五、本清单维护规则

- 新增对象 → 在「对象清单」加一行 + 建 `skills/<对象>/`、`scripts/<对象>/` 目录。
- 脚本归位 → 从 `store-data-extractor` 复制到 `scripts/<对象>/` 后，在「已有能力」列更新为本库路径，并删除对旧路径的引用。
- 本文件是唯一事实源，脚本 / skill 与清单不一致时，以清单为准并修正脚本。

## 六、脚本归位状态

- [ ] 门店类脚本归位
- [ ] 桌台类脚本归位
- [ ] 群类脚本归位
- [ ] 客户类脚本归位
- [ ] 企微类脚本归位
- [ ] 渠道码类脚本归位
- 归位规则：复制到 `scripts/<对象>/`，按 `对象-动作-说明` 重命名，凭据走 `.env`，写操作补 dryRun。
