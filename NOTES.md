# NOTES — gov-document

## Provenance

- **Source**: ModelScope `llxxxxxxxx/gov-document` (Apache-2.0)
- **Original URL**: https://www.modelscope.cn/skills/llxxxxxxxx/gov-document
- **Fetched**: 2026-08-17 from `https://www.modelscope.cn/skills/llxxxxxxxx/gov-document/archive/zip/master.zip`
- **License**: Apache License 2.0 — see `LICENSE`

## Status

Non-hub third-party skill deployed at user request. Adapted for Hermes Agent at
`~/AppData/Local/hermes/skills/productivity/gov-document/`. Not present in
NousResearch/hermes-agent `skills/` (i.e. NOT a hub skill).

## Changes from upstream (2026-08-17)

This is a list of bug-fixes and feature additions layered on top of the
unmodified ModelScope source. All changes preserve the Apache-2.0 license
and the original JSON schema (additions are backward-compatible).

### HIGH — 4 format types now actually dispatch (was: silently ignored)

**Upstream behavior**: `format_type` (and `doc_type`) were read but the script
produced byte-identical output for `general`/`letter`/`order`/`minutes`/`unknown`/
omitted (all ~37638 bytes).

**Now**: 4 distinct code paths:

| format_type | GB/T 9704 章节 | 关键行为差异 |
|-------------|---------------|-------------|
| `general`   | §3 通用  | 红头机关标志 + 红色版头分隔线 + 发文字号 + 版记 + 4 级层次正文 |
| `letter`    | §4.1 信函 | 机关标志下红色双线(上粗下细) + 下页边红色双线(上细下粗) + 无发文字号 + 无版记 + 首页不显示页码 |
| `order`     | §4.2 命令(令) | 机关标志由全称+「命令」/「令」 + 令号居中独占 + 上距版心 20mm + 无主送机关 |
| `minutes`   | §4.3 纪要 | "×××××纪要"标志 + 出席人员名单(「出席:」后三号仿宋) |

Plus unknown values fall back to `general` with a stderr warning.

### MEDIUM — page numbers per GB/T 9704 §3.4

**Upstream behavior**: single center-aligned footer paragraph with PAGE field.

**Now**:
- `evenAndOddHeaders` enabled in `settings.xml`
- 奇数页 footer: `—PAGE—`, alignment `RIGHT`
- 偶数页 footer: `—PAGE—`, alignment `LEFT`
- 一字线 (`—`) 字号 14pt (四号半角宋体 — 实际用了「宋体」作为东亚字体)

### LOW additions

| Fix | Description |
|-----|-------------|
| 1 | Bracket normalization in `format_doc_number(prefix, year, sequence)`: ASCII `[]`, 全角 `[]`, 六角 `〔〕` 都归一为六角 `〔〕`. 用户写 `2026` / `[2026]` / `［2026］` 都会得到 `〔2026〕`. |
| 2 | Attachments use explicit `<w:br/>` via `run.add_break()` instead of relying on python-docx's implicit `'\n'` → `<w:br/>` conversion. |
| 3 | `issuer_names` 数组 (联合行文): 每个机关一行右对齐。`issuer_name` 仍受支持；都给则 `issuer_names` 优先。 |
| 4 | `title_lines` 数组 (梯形/菱形排版): 每行一个 paragraph 居中。`title` 仍受支持；都给则 `title_lines` 优先。 |
| 5 | `order_number` (命令(令) 格式专用): 居中独占编排于机关标志下空二行。 |
| 6 | `attendees` (纪要格式专用): 字符串或数组，在主体 build_body 之后追加。 |
| 7 | Hygiene files: `LICENSE`, `.gitignore`. |

## Schema additions (backward-compatible)

```diff
  "subject": {
    "title": "...",                  // 单行标题 (existing)
+   "title_lines": ["...", "..."],   // 梯形/菱形排版 (NEW)
    "recipients": [...],
    "body": [...],
    "issuer_name": "...",            // 单署名 (existing)
+   "issuer_names": ["...", "..."],  // 联合行文 (NEW)
    "issue_date": "...",
+   "attendees": "...",              // 纪要专用 (NEW)
  }

  "header": {
    ...
+   "order_number": "..."            // 命令(令)专用 (NEW)
  }
```

Existing inputs that omit these new fields keep working unchanged.

## Known limitations

1. **印章 (seal) 占位**: `seal: true` is read but no image is embedded — the
   script just sets `RIGHT` alignment and trusts the user to add the actual
   印章 (PNG/JPG) later. Adding `add_picture()` with a 印章 overlay is the
   next step.
2. **Page-count limit**: `WD_BREAK.PAGE` is not used — multi-page docs rely on
   Word's auto-pagination. For very long body content (>1 page) this works
   fine; for >2 pages the odd/even footer switch is verified.
3. **No PDF conversion**: the upstream README mentioned "可选转换：将 docx 转
   换为 pdf" but no PDF script was shipped. Use `soffice --headless --convert-to pdf`
   or `docx2pdf` separately if needed.
4. **纪要 doc_number**: §4.3 doesn't strictly require 发文字号, but the
   minutes dispatch in `build_header` doesn't currently suppress
   `doc_number` if the user provides one. Add `if format_type == 'minutes':
   doc_number = {}` if that becomes an issue.

## Verification

Reproduced 2026-08-17 with:
- Python 3.10.8 (`C:/Users/Jett_/AppData/Local/Programs/Python/Python310/python.exe`)
- `python-docx 1.2.0` (installed via `pip install python-docx` to Python310)
- Probe fixtures generated in temp, results: 4/4 formats distinct,
  bracket normalization correct, attachments use `<w:br/>`, signatures
  split per 联合行文 member, page footers per odd/even + 一字线.

## Cross-skill impact

None. This skill is self-contained and does not reference any other skill
or hub tool.