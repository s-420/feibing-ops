# skill 库

> 按「对象」分目录。每个子目录放该对象下的 skill（SKILL.md 形式）。

## 目录规则

- 一级目录 = 对象名（与 `taxonomy/index.md` 对象清单一致）。
- 每个 skill 一个子目录，内含 `SKILL.md`。

## skill 是什么 / 何时用

- skill = 给 agent 的工作流程说明（「遇到什么 → 怎么做 → 检查什么 → 异常怎么办」）。
- 用于「需要判断、多步组合」的场景，skill 内部可调用 `scripts/` 里的脚本。

## 新增一个 skill

1. 复制 `templates/skill模板.md`。
2. 填元数据头（对象 / 动作 / 风险 / 输入 / 输出 / 验收）。
3. 按模板写工作流。
4. 在 `taxonomy/index.md` 登记。
