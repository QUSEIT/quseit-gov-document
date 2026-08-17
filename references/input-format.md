# JSON输入格式说明

生成公文时，需要提供一个JSON文件作为输入，格式如下：

## 文种与 format_type 对应（依据《党政机关公文处理工作条例》第八条）

《条例》第八条规定公文种类共 15 种。本技能**全部支持**，但其中只有 3 种有专门版式分支，其余 12 种统一走通用版式（`general`）：

| 文种 | format_type | 专门版式 |
|------|------------|---------|
| 决议 | `general` | 否 |
| 决定 | `general` | 否 |
| 命令(令) | `order` | **是**（§4.2 — 机关标志由全称+「命令」/「令」、独占编排令号、上距版心 20mm）|
| 公报 | `general` | 否 |
| 公告 | `general` | 否 |
| 通告 | `general` | 否 |
| 意见 | `general` | 否 |
| 通知 | `general` | 否 |
| 通报 | `general` | 否 |
| 报告 | `general` | 否 |
| 请示 | `general` | 否 |
| 批复 | `general` | 否 |
| 议案 | `general` | 否 |
| 函 | `letter` | **是**（§4.1 — 红色双线上下、无发文字号、无版记、首页不显示页码）|
| 纪要 | `minutes` | **是**（§4.3 — "×××××纪要"标志、出席人员名单）|

> **JSON 写法**：JSON 顶层字段仍是 `format_type`（不直接接受文种名称）。如要起草"决议"，在 JSON 里写 `"format_type": "general"` 并把标题写为"关于×××的决议"。
> **缺失字段**：如果你忘记填 `format_type`，脚本会按 `general` 处理并打印 `format_type='X' 不在 ... 内，按 general 处理` 警告到 stderr。

## 完整JSON结构

```json
{
  "doc_type": "general",
  "paper_size": "A4",
  "header": {
    "copy_number": "000001",
    "security_level": "秘密",
    "security_period": "5",
    "urgency": "特急",
    "issuer_mark": "国务院文件",
    "doc_number": {
      "prefix": "国发",
      "year": "2026",
      "sequence": "1"
    },
    "signer": ["张三", "李四"]
  },
  "subject": {
    "title": "关于XXX的通知",
    "recipients": ["各省、自治区、直辖市人民政府", "国务院各部委"],
    "body": [
      "正文第一段内容...",
      "正文第二段内容...",
      {
        "level": 1,
        "text": "一级标题",
        "paragraphs": ["该标题下的段落..."]
      },
      {
        "level": 2,
        "text": "（一）二级标题",
        "paragraphs": ["该标题下的段落..."]
      }
    ],
    "attachments": [
      {"name": "附件一名称", "file": "path/to/file1.docx"},
      {"name": "附件二名称"}
    ],
    "issuer_name": "国务院",
    "issue_date": "2026年8月12日",
    "seal": true,
    "notes": "（此件公开发布）"
  },
  "footer": {
    "cc": ["国务院办公厅", "国家发展改革委"],
    "print_org": "国务院办公厅秘书局",
    "print_date": "2026年8月12日"
  },
  "page_number": true,
  "format_type": "general"
}
```

## 字段说明

### 顶层字段

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `doc_type` | string | 否 | 公文类型: general(通用)/letter(信函)/order(命令)/minutes(纪要)，默认 general。**已废弃**，推荐用 `format_type` |
| `paper_size` | string | 否 | 纸张尺寸，默认 A4 |
| `header` | object | 是 | 版头要素 |
| `subject` | object | 是 | 主体要素 |
| `footer` | object | 否 | 版记要素 |
| `page_number` | boolean | 否 | 是否显示页码，默认 true |
| `format_type` | string | 否 | 格式类型: general/letter/order/minutes，默认 general |

> **format_type 与公文格式对应表**:
>
> | format_type | GB/T 9704 章节 | 关键特征 |
> |-------------|---------------|---------|
> | `general`   | §3 通用 | 红头机关标志(26pt 小标宋红色)、红色版头分隔线、发文字号、4 级层次正文、版记抄送+印发 |
> | `letter`    | §4.1 信函 | 机关标志下 4mm 红色双线(上粗下细)、下页边 20mm 红色双线(上细下粗)、无发文字号、无版记、第一页不显示页码 |
> | `order`     | §4.2 命令(令) | 机关标志由全称+「命令」/「令」、独占编排令号、上距版心 20mm |
> | `minutes`   | §4.3 纪要 | "×××××纪要"标志、上距版心 35mm、出席人员名单 |

### header 对象

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `copy_number` | string | 否 | 份号，6位数字 |
| `security_level` | string | 否 | 密级: 秘密/机密/绝密 |
| `security_period` | string | 否 | 保密期限(年) |
| `urgency` | string | 否 | 紧急程度: 急件/特急 |
| `issuer_mark` | string | 是 | 发文机关标志，如"国务院文件" |
| `doc_number` | object | 是 | 发文字号 |
| `doc_number.prefix` | string | 否* | 机关代字，如"国发" |
| `doc_number.year` | string | 否* | 年份（任意括号 ASCII []/全角 []/六角 〔〕 都会被归一为六角） |
| `doc_number.sequence` | string | 否* | 顺序号 |
| `signer` | array | 否 | 签发人姓名列表（上行文布局：发文字号居左，签发人居右） |
| `order_number` | string | 否 | **命令(令)格式专用** — 令号（如"第1号"），居中独占编排 |

* `doc_number` 在 `general` 格式下必填；在 `letter`/`order` 格式下省略。

### subject 对象

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `title` | string | 否* | 公文标题（单行）；与 `title_lines` 二选一 |
| `title_lines` | array | 否* | 标题分行（梯形/菱形排版）；每个元素为一行 |
| `recipients` | array | 是 | 主送机关列表 |
| `body` | array | 是 | 正文内容(字符串或结构化段落) |
| `attachments` | array | 否 | 附件列表 |
| `issuer_name` | string | 否* | 发文机关署名（单署名）；与 `issuer_names` 二选一 |
| `issuer_names` | array | 否* | 联合行文署名列表（主办机关在前）；与 `issuer_name` 二选一 |
| `issue_date` | string | 是 | 成文日期，格式"YYYY年M月D日" |
| `seal` | boolean | 否 | 是否加盖印章，默认 true |
| `notes` | string | 否 | 附注内容 |
| `attendees` | string/array | 否 | **纪要格式专用** — 出席人员名单（"出席"二字后由全角冒号引导，姓名用三号仿宋） |

* `title` 与 `title_lines` 二选一；`issuer_name` 与 `issuer_names` 二选一；都给则 `title_lines` 优先、`issuer_names` 优先。

### body 数组元素

正文可以是简单字符串数组，也可以包含结构化层次:

**简单格式:**
```json
["第一段内容...", "第二段内容..."]
```

**结构化格式:**
```json
[
  "普通段落...",
  {"level": 1, "text": "一、一级标题(黑体)", "paragraphs": ["段落..."]},
  {"level": 2, "text": "（一）二级标题(楷体)", "paragraphs": ["段落..."]},
  {"level": 3, "text": "1.三级标题(仿宋)", "paragraphs": ["段落..."]},
  {"level": 4, "text": "（1）四级标题(仿宋)", "paragraphs": ["段落..."]}
]
```

### footer 对象

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `cc` | array | 否 | 抄送机关列表 |
| `print_org` | string | 否 | 印发机关 |
| `print_date` | string | 否 | 印发日期 |

## 简化示例（最简输入）

```json
{
  "header": {
    "issuer_mark": "XX市人民政府文件",
    "doc_number": {
      "prefix": "X政发",
      "year": "2026",
      "sequence": "1"
    }
  },
  "subject": {
    "title": "关于做好2026年度工作的通知",
    "recipients": ["各县区人民政府"],
    "body": ["正文内容..."],
    "issuer_name": "XX市人民政府",
    "issue_date": "2026年8月12日"
  }
}
```

## 各格式示例

### 1. 命令(令)格式（order）

```json
{
  "header": {
    "issuer_mark": "中华人民共和国主席令",
    "order_number": "第八十一号"
  },
  "subject": {
    "body": [
      "为表彰在科学技术进步活动中作出突出贡献的个人和组织，根据《国家科学技术奖励条例》的规定，现决定：",
      "一、授予×××同志国家最高科学技术奖。",
      "二、授予××项目国家自然科学奖一等奖。"
    ],
    "issuer_name": "中华人民共和国主席",
    "issue_date": "2026年8月12日"
  },
  "format_type": "order"
}
```

### 2. 信函格式（letter）

```json
{
  "header": {
    "issuer_mark": "XX市发展和改革委员会"
  },
  "subject": {
    "title": "关于商请支持我市重点建设项目的函",
    "recipients": ["省发展改革委"],
    "body": ["为推动我市重点项目建设，现就有关事项函告如下：..."],
    "issuer_name": "XX市发展和改革委员会",
    "issue_date": "2026年8月12日"
  },
  "format_type": "letter"
}
```

### 3. 纪要格式（minutes）

```json
{
  "header": {
    "issuer_mark": "XX市城市规划会议纪要"
  },
  "subject": {
    "title": "XX市城市规划委员会2026年第三次会议纪要",
    "body": [
      "会议听取了《XX市国土空间总体规划》编制情况的汇报，审议了...",
      {"level": 1, "text": "一、总体评价", "paragraphs": ["会议认为..."]},
      {"level": 1, "text": "二、工作部署", "paragraphs": ["会议要求..."]}
    ],
    "attendees": ["张三（市长）", "李四（副市长）", "王五（市规划局局长）"],
    "issuer_name": "XX市城市规划委员会办公室",
    "issue_date": "2026年8月12日"
  },
  "format_type": "minutes"
}
```

### 4. 联合行文 + 梯形标题（general）

```json
{
  "header": {
    "issuer_mark": "XX市文化和旅游局 XX市财政局",
    "doc_number": {"prefix": "X文旅联发", "year": "2026", "sequence": "1"}
  },
  "subject": {
    "title_lines": ["XX市文化和旅游局 XX市财政局", "关于印发《XX市文化旅游发展", "专项资金管理办法》的通知"],
    "recipients": ["各县区人民政府"],
    "body": ["为规范文化旅游发展专项资金管理..."],
    "issuer_names": ["XX市文化和旅游局", "XX市财政局"],
    "issue_date": "2026年8月12日"
  },
  "format_type": "general"
}
```
