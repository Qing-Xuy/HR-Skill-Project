# 经验萃取专家 / Experience Extraction Expert

把专家「说得出经历、说不清方法」的隐性经验，整理成别人可以照着做的操作方法、案例和检查单。

An AI skill that turns tacit expert knowledge into reusable playbooks through structured interviews or analysis of existing materials.

## 适用场景

- 访谈业务专家，复盘一个有代表性的成功或失败案例。
- 分析已有的访谈转写稿、会议纪要、聊天记录或工作复盘笔记。
- 将关键判断、实际动作及其证据沉淀为可交接的文档。

## 两种工作方式

| 方式 | 输入 | 处理方式 |
|---|---|---|
| 对话式访谈 | 专家参与对话 | 按五轮访谈协议追问具体事件、判断依据和实施细节，并逐轮请专家纠偏。 |
| 素材式分析 | 已有文本或文件 | 标记事实、推断、关键决策和信息缺口，先给出证据清单与待补问清单。 |

素材读取脚本支持 `.txt`、`.md`、`.docx`、`.srt`、`.vtt`。如果提供的是录音，需要运行环境具备语音转写能力；否则请先提供转写稿。

## 交付物

1. **操作手册与行为复制检查单**：供接手者按步骤执行。
2. **案例总结**：还原情境、关键动作、判断依据和结果证据。
3. **能力与方法价值提炼**：按需生成，例如用于能力盘点或晋升材料。

默认先提供文字版分析框架，待使用者确认或纠偏后再生成文件。输出默认为 Markdown；需要 Word 时可使用仓库内的渲染脚本。

## 开始使用

将整个 [`experience-extraction-expert`](./experience-extraction-expert/) 文件夹复制到支持 `SKILL.md` 的 skills 目录。原始说明以 WorkBuddy 为例：Windows 使用 `%USERPROFILE%\.workbuddy\skills`，macOS / Linux 使用 `~/.workbuddy/skills`。复制后刷新或重启对话。

可以这样发起任务：

> 帮我萃取张工处理客户投诉的经验，请从一个具体案例开始访谈。

> 根据这份访谈转写稿，提炼可复制的方法，并列出还需要补问的问题。

读取 `.docx` 或生成 Word 文档需要 `python-docx`：

```bash
pip install python-docx
```

无需 Word 时可使用 Markdown 输出。更详细的安装和使用说明见 [skill README](./experience-extraction-expert/README.md)。

## 质量原则

- 不编造素材中没有的数字、人名和结果；缺失信息标为「待确认」。
- 区分原始事实与分析推断，让结论能追溯到证据。
- 示例仅展示产出形态，实际结构应由具体案例和素材决定。

## 仓库内容

| 位置 | 内容 |
|---|---|
| [`SKILL.md`](./experience-extraction-expert/SKILL.md) | Skill 的工作流程与质量要求 |
| [`references/`](./experience-extraction-expert/references/) | 访谈协议、提问库、理论、素材分析标准与输出骨架 |
| [`examples/`](./experience-extraction-expert/examples/) | 一份虚构案例示范 |
| [`scripts/`](./experience-extraction-expert/scripts/) | 素材读取、文档渲染与打包脚本 |

Skill 标识符：`experience-extraction-expert`；显示名称：**经验萃取专家 / Experience Extraction Expert**。
