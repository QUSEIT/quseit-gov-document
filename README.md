# gov-document

党政机关公文格式生成器（GB/T 9704-2012）

根据 **GB/T 9704-2012** 国家标准，生成符合规范的党政机关公文（docx）。覆盖《党政机关公文处理工作条例》第八条规定的全部 **15 种公文文种**。

## 功能特性

- ✅ 覆盖全部 15 种公文文种：决议、决定、命令(令)、公报、公告、通告、意见、通知、通报、报告、请示、批复、议案、函、纪要
- ✅ 4 种格式版式：通用（general）、信函（letter）、命令(令)（order）、纪要（minutes），分别对应 GB/T 9704-2012 第 3 章通用格式和第 4 章特殊格式
- ✅ A4 页面、天头 37mm、订口 28mm、版心 156mm×225mm
- ✅ 三号仿宋正文、二号小标宋标题、黑体一级层次、楷体二级层次
- ✅ 红色发文机关标志（红头）与红色版头分隔线
- ✅ 发文字号六角括号归一化（`[2026]` / `［2026］` / `〔2026〕` 均归一为 `〔2026〕`）
- ✅ 奇偶页页码：单页码居右、双页码居左，带一字线（—PAGE—）
- ✅ 联合行文（`issuer_names` 数组，主办机关在前）
- ✅ 梯形/菱形标题排版（`title_lines` 数组）
- ✅ 密级与保密期限、紧急程度、份号、签发人（上行文版式）
- ✅ 附件说明（`<w:br/>` 分行）、附注、抄送机关、印发机关与印发日期

## 安装

### 依赖

```bash
pip install python-docx>=1.0.0
```

### 作为 Hermes 技能安装

将本仓库拷贝到 Hermes 技能目录：

```bash
cp -r quseit-gov-document "$HERMES_HOME/skills/productivity/gov-document"
```

Hermes 会通过 `SKILL.md` 自动识别。当用户提到"红头文件 / 公文 / 通知 / 请示 / 批复 / 函 / 纪要"等关键词时自动触发。

## 使用方法

### 方式一：命令行脚本

```bash
python scripts/generate_gov_docx.py --input input.json --output output.docx
```

### 方式二：在 Hermes 对话中

> 请帮我起草一份关于 2026 年端午节放假安排的通知，用上海市人民政府红头文件。

Hermes 会按工作流收集要素、生成 JSON、调用脚本产出 docx。

### JSON 输入示例（通用公文）

```json
{
  "header": {
    "issuer_mark": "XX市人民政府文件",
    "doc_number": {"prefix": "X政发", "year": "2026", "sequence": "1"}
  },
  "subject": {
    "title": "关于做好2026年度工作的通知",
    "recipients": ["各县区人民政府"],
    "body": ["为做好2026年度工作，现将有关事项通知如下。"],
    "issuer_name": "XX市人民政府",
    "issue_date": "2026年8月17日"
  },
  "footer": {
    "cc": ["市委办公室"],
    "print_org": "XX市人民政府办公室",
    "print_date": "2026年8月17日"
  },
  "format_type": "general"
}
```

### 文种 → format_type 映射

| 文种 | format_type | 专门版式 |
|------|------------|---------|
| 决议、决定、公报、公告、通告、意见、通知、通报、报告、请示、批复、议案 | `general` | 否 |
| 命令(令) | `order` | 是（§4.2）|
| 函 | `letter` | 是（§4.1）|
| 纪要 | `minutes` | 是（§4.3）|

完整字段说明见 [`references/input-format.md`](references/input-format.md)。

## 文件结构

```
gov-document/
├── SKILL.md                     # Hermes 技能入口
├── scripts/
│   └── generate_gov_docx.py     # 公文生成脚本
├── references/
│   ├── format-spec.md           # GB/T 9704-2012 完整规范参考
│   └── input-format.md          # JSON 输入格式说明 + 各格式示例
├── requirements.txt             # python-docx>=1.0.0
├── LICENSE                      # Apache-2.0
├── NOTES.md                     # 变更日志与已知问题
└── README.md                    # 本文档
```

## 许可证

[Apache License 2.0](LICENSE)

原技能由 ModelScope 用户 `llxxxxxxxx` 发布（Apache-2.0），本仓库为 Hermes 适配版，包含格式分支实现、页码修复、联合行文等增强。
