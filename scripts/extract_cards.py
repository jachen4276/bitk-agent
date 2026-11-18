# -*- coding: utf-8 -*-
"""
从 丹寨苗族蜡染-知识卡片数据集.html 抽取 102 张知识卡片为结构化 JSON。

用途：为「可溯源问答 Agent」构建检索用的知识库底座。
输入：丹寨苗族蜡染-知识卡片数据集.html（团队原始材料，只读）
输出：data/cards.json、data/cards.csv、data/cards_stats.md

设计原则：
  1. 只转录原文，不做任何改写、补全、合并；
  2. 每条卡片保留 来源编号 / 来源标题 / 来源URL / 原文摘录，保证可回溯；
  3. 解析不到的字段留空字符串，绝不推测填写。
"""
import csv
import json
import re
import sys
from pathlib import Path

SRC = Path(
    r"C:\Users\jachen\Desktop\JustChange\比赛等\创挑赛\正式比赛\丹寨苗族蜡染-知识卡片数据集.html"
)
OUT_DIR = Path(__file__).resolve().parent.parent / "data"


def strip_tags(s: str) -> str:
    """去掉标签并还原实体，保留纯文本。"""
    s = re.sub(r"(?s)<[^>]+>", "", s)
    s = (s.replace("&nbsp;", " ").replace("&amp;", "&")
          .replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"'))
    return re.sub(r"\s+", " ", s).strip()


def parse(html_text: str):
    blocks = re.findall(r'(?s)<article class="card[^"]*" id="(K\d{3})">(.*?)</article>', html_text)
    cards = []
    for kid, body in blocks:
        rec = {
            "id": kid,
            "类别": "",
            "标题": "",
            "正文": "",
            "原文摘录": "",
            "来源编号": "",
            "来源标题": "",
            "来源URL": "",
            "可信等级": "",
            "关键词": [],
            "关联卡片": [],
        }

        m = re.search(r'(?s)<span class="cat">(.*?)</span>', body)
        if m:
            rec["类别"] = strip_tags(m.group(1))

        m = re.search(r"(?s)<h4>(.*?)</h4>", body)
        if m:
            rec["标题"] = strip_tags(m.group(1))

        m = re.search(r'(?s)<div class="body">(.*?)</div>', body)
        if m:
            rec["正文"] = strip_tags(m.group(1))

        m = re.search(r'(?s)<blockquote class="quote">(.*?)</blockquote>', body)
        if m:
            q = strip_tags(m.group(1))
            rec["原文摘录"] = re.sub(r"^原文摘录\s*", "", q)

        # 来源：S023 · <a href="URL">标题</a>
        m = re.search(
            r'(?s)来源\s*(S\d{3})\s*·\s*<a href="([^"]+)"[^>]*>(.*?)</a>', body
        )
        if m:
            rec["来源编号"] = m.group(1)
            rec["来源URL"] = m.group(2)
            rec["来源标题"] = strip_tags(m.group(3))

        m = re.search(r'(?s)<span class="trust">(.*?)</span>', body)
        if m:
            rec["可信等级"] = strip_tags(m.group(1))

        rec["关键词"] = [strip_tags(x) for x in re.findall(r'(?s)<span class="kw">(.*?)</span>', body)]

        for rid, rtitle in re.findall(
            r'(?s)<a href="#(K\d{3})"[^>]*title="([^"]*)"', body
        ):
            rec["关联卡片"].append({"id": rid, "title": rtitle})

        cards.append(rec)
    return cards


def main() -> int:
    if not SRC.exists():
        print(f"[ERROR] 找不到源文件: {SRC}")
        return 1

    cards = parse(SRC.read_bytes().decode("utf-8"))
    if len(cards) != 102:
        print(f"[WARN] 解析到 {len(cards)} 张卡片，与预期的 102 张不符，请检查源文件结构。")

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    id2title = {c["id"]: c["标题"] for c in cards}

    # JSON：检索/评测直接用
    (OUT_DIR / "cards.json").write_text(
        json.dumps(cards, ensure_ascii=False, indent=1), encoding="utf-8"
    )

    # CSV：人工核对/Excel 打开用
    cols = ["id", "类别", "标题", "正文", "原文摘录", "来源编号", "来源标题", "来源URL", "可信等级"]
    with (OUT_DIR / "cards.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(cols + ["关键词", "关联卡片"])
        for c in cards:
            w.writerow([c[k] for k in cols] + ["|".join(c["关键词"]), "|".join(x["id"] for x in c["关联卡片"])])

    # 质量自查：这些数字决定后面的实验能不能做
    from collections import Counter

    cat = Counter(c["类别"] for c in cards)
    src = Counter(c["来源编号"] for c in cards)
    no_quote = [c["id"] for c in cards if not c["原文摘录"]]
    no_url = [c["id"] for c in cards if not c["来源URL"]]
    bad_rel = [(c["id"], r["id"]) for c in cards for r in c["关联卡片"] if r["id"] not in id2title]
    self_rel = [c["id"] for c in cards if any(r["id"] == c["id"] for r in c["关联卡片"])]
    rel_cnt = sum(len(c["关联卡片"]) for c in cards)
    dup_title = [t for t, n in Counter(c["标题"] for c in cards).items() if n > 1]

    lines = [
        "# cards.json 质量自查（脚本生成，勿手改）",
        "",
        f"- 卡片总数：**{len(cards)}**",
        f"- 来源文章数：**{len(src)}**",
        f"- 类别数：**{len(cat)}**",
        f"- 关键词总数：**{sum(len(c['关键词']) for c in cards)}**",
        f"- 关联卡片边数：**{rel_cnt}**（平均 {rel_cnt / len(cards):.1f} 条/卡）",
        f"- 正文总字数：**{sum(len(c['正文']) for c in cards)}**",
        "",
        "## 类别分布",
        "",
    ]
    for k, v in cat.most_common():
        lines.append(f"- {k}：{v}")
    lines += ["", "## 完整性自查", ""]
    lines.append(f"- 缺原文摘录：{len(no_quote)} 张 {no_quote if no_quote else ''}")
    lines.append(f"- 缺来源URL：{len(no_url)} 张 {no_url if no_url else ''}")
    lines.append(f"- 关联卡片指向不存在的ID：{len(bad_rel)} 处 {bad_rel if bad_rel else ''}")
    lines.append(f"- 关联卡片指向自己：{len(self_rel)} 张 {self_rel if self_rel else ''}")
    lines.append(f"- 标题重复：{len(dup_title)} 组 {dup_title if dup_title else ''}")
    lines += ["", "## 与「纹样」直接相关的卡片（决定 Agent 能回答什么）", ""]
    for c in cards:
        if c["类别"] in ("纹样图案", "工艺原理"):
            lines.append(f"- {c['id']} [{c['类别']}] {c['标题']}")
    (OUT_DIR / "cards_stats.md").write_text("\n".join(lines), encoding="utf-8")

    print(f"[OK] {len(cards)} 张卡片 -> {OUT_DIR / 'cards.json'}")
    print(f"[OK] 自查报告 -> {OUT_DIR / 'cards_stats.md'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
