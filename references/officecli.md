# officecli pptx 速查

本 skill 用到的命令语法，按 v1.0.152 实测。规则与陷阱在 SKILL.md（§3.3 公式、§4 克隆、§6 陷阱），
这里只放写法。属性名不在这里的，查 `officecli help pptx <element>`（加 `--json` 出结构化 schema），不要猜。

## 路径

- 页与形状：`'/slide[3]'`、`'/slide[3]/shape[@id=31]'`、`'/slide[3]/picture[@id=130]'`、`'/slide[3]/notes'`；
  段落与 run：`.../shape[@id=5]/paragraph[1]/run[5]`。下标从 1 开始。
- 路径一律加引号，否则 shell 会把 `[N]` 当通配符展开。
- 优先用 `get` 返回的 `@id=` 路径：位置下标会随增删漂移，id 不会（但克隆页会重新分配 id）。

## 看结构

```bash
officecli get deck.pptx '/slide[3]' --depth 1          # 列出该页所有形状：路径、名字、几何、文字预览
officecli get deck.pptx '/slide[3]' --depth 4          # 展开到段落和 run，查哪一段是哪个 run
officecli query deck.pptx 'picture' --json             # 全 deck 按类型查；支持 [attr=value]、:contains("文字")
```

## 改文字与格式

```bash
# 替换：保留原 run 格式；路径决定范围，写到页或形状，不要写 /
officecli set deck.pptx '/slide[4]' --find '<样板句>' --replace '<新句>'
# 给匹配文字设格式：自动拆 run，可跨 run 匹配
officecli set deck.pptx '/slide[7]' --find 'a(?=∥)' --prop regex=true --prop italic=true --prop 'font=Times New Roman'
officecli set deck.pptx '/slide[7]' --find '∥' --prop subscript=true
# 整个形状或 run 设属性
officecli set deck.pptx '/slide[7]/shape[@id=5]' --prop align=left
```

无匹配也返回成功；加 `--json` 看 `"matched": N` 确认真的改到了。区分大小写，不区分用 `--prop 'find=(?i)...'`。

## 形状与图片

```bash
# 插图：只给 width 时按原图宽高比算高度；长度单位 in / cm / pt / emu
officecli add deck.pptx '/slide[3]' --type picture --prop src=fig/a-trim.png --prop x=6.9in --prop y=2.7in --prop width=5in
# 新建中文文本框：lang=zh-CN 才按中文避头尾断行（默认 en-US）
officecli add deck.pptx '/slide[3]' --type textbox --prop text='图注' --prop font=微软雅黑 --prop size=14 --prop lang=zh-CN --prop x=1in --prop y=6in --prop width=4in --prop height=0.4in
# 挪位置 / 改尺寸（面板、旗标、图片同一写法）
officecli set deck.pptx '/slide[3]/shape[@id=8]' --prop x=0.38in --prop y=2.43in --prop width=5.2in --prop height=4.3in
# 删形状（如版式 6 的页底 ✓ 行）
officecli remove deck.pptx '/slide[3]/shape[@id=31]'
# 复制页 / 复制形状：带上全部关联的部件
officecli add deck.pptx / --from '/slide[6]'
```

## 演讲者备注

```bash
officecli add deck.pptx '/slide[3]' --type notes --prop text='口播稿……'   # 该页还没有备注时
officecli set deck.pptx '/slide[3]/notes' --prop text='口播稿……'           # 覆盖已有备注
officecli get deck.pptx '/slide[3]/notes'
```

## 批处理

整份文案写进一个 JSON，一次执行。**原子执行**：任一条失败则整批回滚，文件保持原样；
返回的 JSON 里逐条列出成败。字段与命令行一一对应：`command`、`path`、`parent`、`type`、`from`、`props`。

```json
[
  {"command": "set", "path": "/slide[4]", "props": {"find": "<样板句>", "replace": "<新断言句>"}},
  {"command": "remove", "path": "/slide[4]/shape[@id=31]"},
  {"command": "add", "parent": "/slide[4]", "type": "notes", "props": {"text": "口播稿……"}}
]
```

```bash
officecli batch deck.pptx --input fill.json --json
```

## 原始 XML（DOM 命令做不到时才用）

```bash
officecli raw deck.pptx '/slide[6]'                                          # 看该页 XML
officecli raw-set deck.pptx '/slide[6]' --xpath '//p:timing' --action remove  # 删动画节点
```

`--action` 可选 `append / prepend / insertbefore / insertafter / replace / remove / setattr`；不用写 xmlns 声明。

## 收尾

```bash
officecli close deck.pptx       # 落盘并释放文件；python-pptx、COM、PowerPoint 读文件前必须先 close
officecli validate deck.pptx    # 按 OpenXML schema 校验
```

## Shell 注意

- 文字里有 `$` 用单引号：`--prop text='$15M'`，双引号会把 `$15` 当变量吃掉。
- 值里有空格要整体加引号：`--prop 'font=Times New Roman'`。
- 文字里的换行写 `\\n`；多段文案更稳的做法是走 batch JSON。
