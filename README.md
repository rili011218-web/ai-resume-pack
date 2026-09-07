# AI Resume Pack · AI 简历制作技能包

> 让 AI 帮你做出一份**能投出去、经得起面试追问**的中文简历。
> A skill pack that helps you build a job-ready resume with AI — tailored, truthful, and interview-proof.

## 这是什么 / What is this

一份开源、免费、**不含任何个人信息**的简历制作工具包。无论你想投**投资、咨询、互联网、快消还是其他方向**，都可以按下面的流程，让任意 AI 助手（DeepSeek、Claude、GPT 等）陪你从零产出一份可投递的简历：

1. **填问卷**（`questionnaire.md`）—— 一次性收集齐简历素材
2. **让 AI 按 SKILL.md 的方法做**—— 把问卷结果 + 本仓库的 `SKILL.md` 给 AI
3. **用生成器出 Word**（`generator/`）—— 填好配置，一键生成排版好的 .docx
4. **看定制指南 + 面试指南**（`guides/`）—— 一个方向一版简历，每条经历都能讲

## 仓库内容 / Contents

| 文件 | 用途 |
|---|---|
| `questionnaire.md` | 📋 **标准问卷**：用户照着填，把答案粘贴给 AI 即开工 |
| `SKILL.md` | 🧠 **简历生产方法论**：让 AI 加载执行的"技能文件"（兼容 Claude Code 技能格式） |
| `guides/multi-direction-customization.md` | 🎯 **多方向定制指南**：一份主简历 → 按方向出多个版本 |
| `guides/interview-prep.md` | 🗣 **面试话术指南**：简历上的每条经历怎么扛住追问 |
| `generator/generate_resume.py` | ⚙️ **简历生成器**：读 config → 一键出 Word（python-docx） |
| `generator/sample_config.json` | 生成器配置样例（含中文/英文两种示例） |
| `english/english-resume-support.md` | 🌐 **英文简历支持**：英文简历结构、常用动词与表达 |

## 快速开始 / Quickstart

### 方式 A：纯对话（最快，不需要装任何东西）

1. 打开 `questionnaire.md`，按里面的问题填好你的信息（打字即可，不用一次填完）
2. 把填好的内容 + 这句话发给任意 AI：
   > "请阅读并遵循这个仓库的 `SKILL.md` 和 `questionnaire.md`，帮我做一份投【方向】的简历。"
3. AI 会补齐提问 → 产出简历文字稿 → 你再决定是否用生成器排版

### 方式 B：本地生成 Word（需要 Python 3 + python-docx）

```bash
cd generator
pip install python-docx
cp sample_config.json my_config.json   # 把 my_config.json 改成你的信息
python3 generate_resume.py my_config.json   # 生成 .docx
```

生成器字段说明见 `generator/sample_config.json` 内的示例与 `generator/README.md`。

## 核心原则 / Principles

- **简历美化 ≠ 简历造假**：只做强措辞、结构调整、亮点前置；编造经历、数字、成果是红线——简历上的每个字都要经得起面试追问与背景调查
- **一稿多投要定制**：投不同方向用不同版本（排序、措辞、技能侧重点都不同），别一份简历打天下
- **会写也要会讲**：简历是面试的"提纲"，每条经历都要准备 1 分钟能讲完的故事

## 贡献 / Contribution

欢迎 PR：补充更多方向的定制示例、面试高频题、英文词库、其他语言的生成器支持。新增内容请确保**不含个人信息**。

## License

MIT — 详见 [LICENSE](LICENSE)。

---

*Made with ❤️ for job seekers. 所有示例内容均为占位符，可放心公开使用。*
