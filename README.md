# p-ppt-academic

给 Claude Code / Codex 用的学术汇报 PPT skill：从零做答辩、组会、开题、项目汇报的 `.pptx`。

视觉基准来自一份浙大应用力学所的博士答辩 deck，叙事基准是 Assertion–Evidence 的中式答辩变体
——标题栏只做导航，断言句另立横幅，证据占中部 ≥60% 版面。

## 安装

仓库里已经装好两份，内容相同：

| Agent | 路径 |
|---|---|
| Claude Code | `.claude/skills/p-ppt-academic/` |
| Codex | `.agents/skills/p-ppt-academic/` |

装到别的仓库：把其中一份整目录复制过去即可。装完**重启会话**才会被识别。

## 内容

```
p-ppt-academic/
├── SKILL.md                    # 主流程：叙事骨架 → 版式计划 → 克隆 → 填内容 → 质检门
├── assets/reference-deck.pptx  # 10 页版式样板，克隆源
├── scripts/clone_pages.ps1     # PowerPoint COM 按版式计划克隆页面
├── scripts/export_png.ps1      # COM 导出逐页 PNG 供目检
├── scripts/qa_check.py         # 越界 / 溢出 / 样板残留词 / 数字与语料交叉核验
└── agents/openai.yaml          # Codex 接口声明
```

## 依赖

- Windows + PowerPoint（COM 自动化，`New-Object -ComObject PowerPoint.Application`）
- Python + `python-pptx`

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
