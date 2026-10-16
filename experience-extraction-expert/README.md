# 经验萃取专家 / Experience Extraction Expert · 安装与使用

把散落在个人身上的隐性经验，萃取成可被他人照抄的方法与文档。

---

## 安装

1. 把整个 `experience-extraction-expert` 文件夹，复制到本机的 skills 目录：

   | 系统 | 路径 |
   |---|---|
   | Windows | `%USERPROFILE%\.workbuddy\skills` |
   | macOS / Linux | `~/.workbuddy/skills` |

2. 重启或刷新一次对话，即可通过以下方式触发：

   > 帮我萃取一下 XX 的经验
   > 把这段访谈录音提炼成方法论
   > experience extraction / tacit knowledge extraction

---

## 可选依赖

**核心功能零依赖。** 只有生成 Word 文档时才需要 `python-docx`：

```bash
pip install python-docx
```

装不上也没关系——Skill 会自动降级为 Markdown 输出，功能不受影响。

推荐装了 `python-docx` 的完整体验：

- ✅ 生成排版规范的 Word 三件套
- ❌ 未安装 → 自动输出 Markdown 三件套

---

## 两种用法

### 模式 A｜对话式（专家本人在场）

```
帮我萃取一下张工处理客户投诉的经验
```

Skill 会扮演萃取者，用五轮协议逐步深挖，过程中不断回显它"听到的事实"和"盯住的隐性判断"，请专家纠偏。

### 模式 B｜素材式（已有转写稿 / 录音文本）

```
（给出文件路径）按经验萃取处理一下
```

支持格式：`.txt` `.md` `.docx` `.srt` `.vtt`

**录音怎么办？**

- 如果本机配置语音转写能力（如云端 ASR Key），Skill 会尝试自动转写。
- 如果不可用，Skill 会明确提示，请先转成文本再给它。**不会卡住，也不会假装听懂。**

---

## 产出

默认节奏：**先给文字版框架 → 你确认或纠偏 → 再生成文件。**

| 文档 | 用途 |
|---|---|
| ① 操作手册 + 行为复制检查单 | 给接手的人，照单执行即达基础质量 |
| ② 案例总结 | 给新人，看懂这件事当时是怎么发生的 |
| ③ 能力与方法价值提炼 ⚙️ | 可选项，需要时勾选（如"还要一份晋升材料"） |

---

## 四条硬规矩

1. **不编造。** 素材没有的数字、时长、结果，一律标"待确认"。
2. **事实与推断分标。** 归纳出来的东西必须注明来源。
3. **不套模板。** `examples/` 里的样例只用于演示产出形态，不代表固定结构。每次按实际素材 + 理论实时推导。
4. **守住范围。** 只做经验萃取。用户离题或受访者跑偏时，Skill 会先接住、再划界、再给出路，请你回到主题——因为萃取靠的是连续追问的深度，一旦打断就退回浅层。

> 一句话判断：**这段对话结束后，会不会变成三件套里的某一行？** 会 → 继续；不会 → 会被礼貌地请出去。

---

## 目录说明

```
experience-extraction-expert/
├── SKILL.md                         主控文件（必读）
├── README.md                        本文件
├── references/
│   ├── extraction-theory.md         理论地图 → 每条理论对应该问哪句话
│   ├── interview-protocol.md        五轮访谈脚本
│   ├── question-bank.md             追问话术库
│   ├── transcript-rubric.md         素材编码表（模式 B 用）
│   └── output-templates.md          三件套骨架
├── examples/
│   └── EX-01-demand-change.md       完整示范（虚构，仅供形态参考）
└── scripts/
    ├── load_source.py               素材读取（txt/md/docx/srt/vtt）
    ├── make_docs.py                 Word / Markdown 渲染器
    └── package_skill.py             打包分发
```

**分发给同事**：运行 `python scripts/package_skill.py`，得到 zip，对方解压到自己 skills 目录即可。
