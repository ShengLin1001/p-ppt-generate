# p-ppt-academic

给 Claude Code / Codex 用的学术汇报 PPT skill：做或改答辩、组会、开题、论坛汇报的 `.pptx`。

视觉基准来自一份浙大应用力学所的博士答辩 deck，叙事基准是 Assertion–Evidence 的中式答辩变体
——标题栏只做导航，断言句另立横幅，证据占中部 ≥60% 版面。

## 安装

仓库根目录就是 skill 本体。实际安装走 [codex-config](https://github.com/ShengLin1001/codex-config)：
它以 git subtree 引入到 `skills-using/project/academic/explicit/p-ppt-academic/`，
由 `pei_ai_univ_reinstall -global -project academic -hook` 装给 Claude Code / Codex。装完**重启会话**才会被识别。

**只能显式调用**：开关只有一处，`agents/openai.yaml` 的 `allow_implicit_invocation: false`（Codex 用 `$p-ppt-academic`）；
Claude Code 侧由 `pei_ai_univ_reinstall -hook` 翻译成 settings.json 的 `skillOverrides: user-invocable-only`（用 `/p-ppt-academic`）。
不在 SKILL.md 加 `disable-model-invocation`。仅适合 Windows（质检导出依赖 PowerPoint COM），已登记在 codex-config 的 `windows-only.txt`。

## 内容

```
p-ppt-generate/                 # 仓库根 = skill 本体
├── SKILL.md                    # 主流程：叙事骨架 → 版式计划 → 克隆 → 填内容 → 质检门
├── assets/reference-deck.pptx  # 10 页版式样板，克隆源
├── scripts/export_png.ps1      # COM 导出逐页 PNG 供目检
├── scripts/qa_check.py         # 越界 / 溢出 / 样板残留词 / 数字交叉核验 / 版面（重叠、填充、空白、上下标）
├── scripts/trim_figure.py      # 插图入页前裁白边
├── references/officecli.md     # officecli pptx 速查：只收本 skill 用到的命令（v1.0.152 实测）
└── agents/openai.yaml          # Codex 接口声明
```

## 依赖

- Windows + PowerPoint（仅质检导出 PNG 用，COM 自动化，`New-Object -ComObject PowerPoint.Application`）
- Python + `python-pptx`
- [officecli](https://github.com/iOfficeAI/OfficeCLI)：克隆页、填内容、修改都用它。升级后用
  `officecli help pptx <element>` 复核 `references/officecli.md` 里的命令，并更新文件头的版本号。
  不要把官方全量手册拉回来覆盖——它九成是 docx/xlsx 内容，还会叫 agent 加载 officecli 自带的设计 skill。

## 样板 deck 的匿名化

`assets/reference-deck.pptx` 的版式取自一份真实的博士答辩 deck，分发前已做匿名化：

| 内容 | 处理 |
|---|---|
| 答辩人、导师姓名 | → 张三 / 李四 |
| 答辩日期 | → 20XX年X月X日 |
| 读博期间发表列表（7 条真实文献） | → 3 条虚构占位条目 |
| 论文题目、目录章节名、四条创新成果 | → 同领域泛化表述 |
| 研究插图 | 已全部删除，只留版式框 |

保留的是断言句、图注与引文行的**写法范例**（它们示范"这类文本该长什么样"），以及机构名与校徽
——这是模板视觉的一部分。deck 里剩下的学科表述均为示例，不对应任何具体个人的研究工作。

## 来源

skill 的规则是四路来源合并后的实测结论（2026-08 在真实答辩 deck 上做过 A/B 比较）：

| 来源 | 取用的部分 |
|---|---|
| [zouchenzhen/thesis-defense-pptx-skill](https://github.com/zouchenzhen/thesis-defense-pptx-skill) | 模板保真路线、COM 渲染质检门 |
| [Gabberflast/academic-pptx-skill](https://github.com/Gabberflast/academic-pptx-skill) | action title、ghost deck test、一页一 exhibit |
| [anthropics/skills](https://github.com/anthropics/skills) `pptx` | python-pptx 编辑陷阱 |
| [wuzhy1ng/research-ppt-review](https://github.com/wuzhy1ng/research-ppt-review) | 中文答辩审查清单、≤3 核心信息/页、量化前置 |
| [PHY041/claude-skill-academic-ppt](https://github.com/PHY041/claude-skill-academic-ppt) | 数字交叉核验、按时长做页数预算 |
| [Michael Alley, Assertion–Evidence](https://www.assertion-evidence.org/) | 断言句 + 视觉证据的单页结构 |

生图路线（整页 AI 出图）一律排除：学术汇报的公式与数据必须可编辑可校对。
