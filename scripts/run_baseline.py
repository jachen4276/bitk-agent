# -*- coding: utf-8 -*-
"""
run_baseline.py —— 三个 baseline 的统一运行脚本（中期检查点用）

【这个脚本干什么】
从 data/cards.json（102 张知识卡片）里检索，回答 data/eval.jsonl（20 道评测题），
自动算分，输出结果表。一次跑一个 baseline。

【三个 baseline 是什么】
  b1  闭卷    不给任何资料，直接让大模型回答。          需要 API
  b2  普通RAG 先检索 3 张最相关卡片塞进提示，再让模型回答。需要 API
  b3  检索直出 不做生成，直接把最相关卡片的原文摘录贴出来。不需要 API

【重要说明：b2 为什么算"普通"RAG】
b2 的提示里已经写了"资料中没有就如实说明"，这是标准 RAG 的普遍做法。
本文档提出的方法要与 b2 拉开差距，靠的不是这句提示，而是：
  ① 强制每条结论绑定 (卡片ID, 原文位置)；② 绑定失败即拒答；
  ③ 生成后做一次证据校验。b2 三者都没有。
这段区别必须写进论文，否则会被问"你这不就是 RAG 吗"。

【怎么跑】
  # 不需要 API，先验证流程通不通（今晚建议先跑这个）
  python scripts/run_baseline.py --method b3

  # 需要 API：Windows PowerShell 里先设环境变量
  $env:OPENAI_API_KEY="sk-xxxx"
  $env:OPENAI_BASE_URL="https://dashscope.aliyuncs.com/compatible-mode/v1"
  $env:MODEL="qwen-plus"
  python scripts/run_baseline.py --method b1
  python scripts/run_baseline.py --method b2

【输出】
  results/{method}_{时间戳}.json   逐题原始结果（写论文/复核都用这个）
  results/{method}_{时间戳}.md     结果表（含指标定义）

【诚实性约束】
  - 本脚本不会编造任何结果：没有 API 就只跑 b3，不会假装跑了 b1/b2；
  - 自动打分是启发式，只做"初判"，论文里的数必须人工复核后才能用；
  - b3 会把原文摘录整段贴出，关键词必然命中，事实正确率会被高估，
    这是 b3 的固有问题，结果表里会自动标注，不能拿它当"最好方法"。
"""
import argparse
import json
import math
import os
import re
import sys
import time
import urllib.error
import urllib.request
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
RESULT_DIR = ROOT / "results"

# 判"是否拒答"的标记词。只是启发式，人工复核为准。
REFUSAL_MARKERS = [
    "资料中没有", "资料未提及", "资料库中没有", "没有记载", "未提及", "没有提及",
    "无法从", "无法回答", "无法确定", "不能确定", "没有找到", "资料不足",
    "未找到", "暂无", "不知道", "无从得知", "无法给出",
]

KID_RE = re.compile(r"K\d{3}")


# ---------------------------------------------------------------- 读取数据
def load_cards():
    path = DATA_DIR / "cards.json"
    if not path.exists():
        sys.exit(f"[ERROR] 找不到 {path}，请先运行 scripts/extract_cards.py")
    return json.loads(path.read_text(encoding="utf-8"))


def load_eval():
    path = DATA_DIR / "eval.jsonl"
    if not path.exists():
        sys.exit(f"[ERROR] 找不到 {path}")
    items = []
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line:
            items.append(json.loads(line))
    return items


# ---------------------------------------------------------------- 检索器
def _norm(s):
    """只保留汉字、字母和数字，去掉标点空白。

    反过来写（保留什么）而不是列举要删什么，避免标点转义出错。
    """
    return re.sub(r"[^\u4e00-\u9fffa-zA-Z0-9]", "", s)


def _bigrams(s):
    s = _norm(s)
    if len(s) <= 1:
        return [s] if s else []
    return [s[i:i + 2] for i in range(len(s) - 1)]


class Retriever:
    """中文字符二元组 + TF-IDF + 余弦相似度。

    为什么不用分词：中文字符二元组不需要词典，对小语料（102 张卡片）够用，
    也少一个依赖。这不是最优检索，但作为 baseline 的检索够公平。
    """

    def __init__(self, cards):
        self.cards = cards
        self.vectors = []
        self.norms = []

        docs = []
        for c in cards:
            text = c["标题"] + "。" + c["正文"] + "。" + "、".join(c["关键词"])
            docs.append(Counter(_bigrams(text)))

        n = len(cards)
        df = Counter()
        for d in docs:
            for g in d:
                df[g] += 1
        self.idf = {g: math.log((n + 1) / (df[g] + 1)) + 1 for g in df}
        self.default_idf = math.log(n + 1) + 1

        for d in docs:
            vec = {g: (1 + math.log(tf)) * self.idf[g] for g, tf in d.items()}
            norm = math.sqrt(sum(v * v for v in vec.values())) or 1.0
            self.vectors.append(vec)
            self.norms.append(norm)

    def search(self, query, k=3):
        q = Counter(_bigrams(query))
        qvec = {g: (1 + math.log(tf)) * self.idf.get(g, self.default_idf) for g, tf in q.items()}
        qnorm = math.sqrt(sum(v * v for v in qvec.values())) or 1.0

        scored = []
        for i, vec in enumerate(self.vectors):
            dot = sum(w * vec.get(g, 0.0) for g, w in qvec.items())
            scored.append((dot / (qnorm * self.norms[i]), i))
        scored.sort(key=lambda x: x[0], reverse=True)
        return [(s, self.cards[i]) for s, i in scored[:k]]


# ---------------------------------------------------------------- 调用大模型
def call_llm(system, user, cfg, retries=3):
    """用 OpenAI 兼容接口调用，不依赖 openai 包（少装一个东西就少一个坑）。"""
    body = {
        "model": cfg["model"],
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
        "temperature": 0,
    }
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        cfg["base_url"].rstrip("/") + "/chat/completions",
        data=data,
        headers={
            "Content-Type": "application/json",
            "Authorization": "Bearer " + cfg["api_key"],
        },
        method="POST",
    )
    last_err = None
    for attempt in range(retries):
        try:
            with urllib.request.urlopen(req, timeout=90) as resp:
                payload = json.loads(resp.read().decode("utf-8"))
            return payload["choices"][0]["message"]["content"].strip()
        except (urllib.error.URLError, KeyError, json.JSONDecodeError) as exc:
            last_err = exc
            time.sleep(2 * (attempt + 1))
    raise RuntimeError(f"调用失败：{last_err}")


def get_api_cfg():
    key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not key:
        return None
    return {
        "api_key": key,
        "base_url": os.environ.get("OPENAI_BASE_URL", "https://dashscope.aliyuncs.com/compatible-mode/v1").strip(),
        "model": os.environ.get("MODEL", "qwen-plus").strip(),
    }


# ---------------------------------------------------------------- 三个方法
def run_b3(item, retriever, topk):
    """检索直出：把 top-1 卡片的原文摘录当答案，不生成。"""
    hits = retriever.search(item["question"], k=topk)
    if not hits:
        return {"answer": "资料中没有找到相关内容。", "cited": []}
    top = [h for h in hits if h[0] > 0.01]
    if not top:
        return {"answer": "资料中没有找到相关内容。", "cited": []}
    quote = top[0][1]["原文摘录"]
    return {"answer": quote, "cited": [c["id"] for _, c in top]}


B1_SYSTEM = "你是一个知识问答助手。请直接回答用户的问题，回答要简洁。"

def run_b1(item, retriever, cfg):
    answer = call_llm(B1_SYSTEM, item["question"], cfg)
    return {"answer": answer, "cited": KID_RE.findall(answer)}


B2_SYSTEM = (
    "你是一个知识问答助手。请根据给定的参考资料回答用户问题。"
    "如果参考资料中没有相关信息，请如实说明。回答要简洁。"
)

def run_b2(item, retriever, cfg, topk):
    hits = retriever.search(item["question"], k=topk)
    ctx = "\n".join(
        f'[{c["id"]}] 标题：{c["标题"]}\n摘录：{c["原文摘录"]}' for _, c in hits
    )
    user = (
        f"【参考资料】\n{ctx}\n\n"
        f"【问题】{item['question']}\n\n"
        "请在回答最后单独一行写明你依据的卡片编号，格式：【引用】K001、K002"
    )
    raw = call_llm(B2_SYSTEM, user, cfg)
    cited = KID_RE.findall(raw)
    answer = re.split(r"【引用】", raw)[0].strip()
    return {"answer": answer, "cited": cited}


# ---------------------------------------------------------------- 打分
def is_refusal(text):
    return any(m in text for m in REFUSAL_MARKERS)


def judge(item, answer, cited):
    """自动初判。返回 (事实是否正确 or None, 是否拒答, 引用命中, 首引命中, 引用精确率)。"""
    refused = is_refusal(answer)

    fact_ok = None
    if item.get("must_any"):
        fact_ok = all(any(w in answer for w in group) for group in item["must_any"])

    gold = set(item.get("gold_cards") or [])
    cite_hit = bool(gold & set(cited)) if gold else None

    # 首引命中：第一张引用的卡片对不对。比"引用里有没有对的"严格得多。
    first_hit = None
    if gold and cited:
        first_hit = cited[0] in gold

    # 引用精确率：引用的卡片里，有多大比例是对的。
    # 这条专治"把所有检索结果全引上"的刷分行为。
    precision = None
    if gold and cited:
        precision = len(gold & set(cited)) / len(cited)

    return fact_ok, refused, cite_hit, first_hit, precision


def score(items, rows):
    """按论文要报的口径统计。指标定义会原样写进结果 md，防止被误读。"""
    fact_items = [r for r in rows if r["type"] in ("answerable", "confusable")]
    ref_items = [r for r in rows if r["type"] == "unanswerable"]

    fact_ok = sum(1 for r in fact_items if r["fact_ok"])
    refused_right = sum(1 for r in ref_items if r["refused"])

    answered_fact = [r for r in fact_items if not r["refused"]]
    with_cite = [r for r in answered_fact if r["cited"]]
    cite_hit = sum(1 for r in with_cite if r["cite_hit"])
    first_hit = sum(1 for r in with_cite if r["first_hit"])
    precisions = [r["cite_precision"] for r in with_cite if r["cite_precision"] is not None]

    # 虚构：该拒答却作答 + 答错且没引用到正确卡片
    halluc = 0
    for r in rows:
        if r["type"] == "unanswerable" and not r["refused"]:
            halluc += 1
        elif r["type"] in ("answerable", "confusable") and r["fact_ok"] is False and not r["cite_hit"]:
            halluc += 1

    n = len(rows)
    return {
        "题目总数": n,
        "事实正确率": f"{fact_ok}/{len(fact_items)} = {fact_ok / len(fact_items) * 100:.1f}%" if fact_items else "—",
        "拒答正确率": f"{refused_right}/{len(ref_items)} = {refused_right / len(ref_items) * 100:.1f}%" if ref_items else "—",
        "首引命中率": f"{first_hit}/{len(with_cite)} = {first_hit / len(with_cite) * 100:.1f}%" if with_cite else "—",
        "引用精确率": f"{sum(precisions) / len(precisions) * 100:.1f}%" if precisions else "—",
        "引用命中率": f"{cite_hit}/{len(with_cite)} = {cite_hit / len(with_cite) * 100:.1f}%" if with_cite else "—",
        "虚构率": f"{halluc}/{n} = {halluc / n * 100:.1f}%",
        "覆盖率": f"{n - sum(1 for r in rows if r['refused'])}/{n} = {(n - sum(1 for r in rows if r['refused'])) / n * 100:.1f}%",
    }


# ---------------------------------------------------------------- 输出
def write_outputs(method, rows, summary, meta):
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = time.strftime("%Y%m%d-%H%M%S")
    base = RESULT_DIR / f"{method}_{stamp}"

    (base.with_suffix(".json")).write_text(
        json.dumps({"meta": meta, "summary": summary, "rows": rows},
                   ensure_ascii=False, indent=1),
        encoding="utf-8",
    )

    lines = [
        f"# baseline 结果：{method.upper()}",
        "",
        f"- 运行时间：{meta['time']}",
        f"- 模型：{meta['model']}",
        f"- 检索 top-k：{meta['topk']}",
        f"- 题目数：{len(rows)}",
        "",
        "## 汇总",
        "",
        "| 指标 | 结果 |",
        "| --- | --- |",
    ]
    for k, v in summary.items():
        lines.append(f"| {k} | {v} |")

    lines += [
        "",
        "## 指标定义（写论文时照抄，别自己改口径）",
        "",
        "- 事实正确率：12 道可答题 + 4 道易混题里，答案包含全部必需关键词的比例；",
        "- 拒答正确率：4 道「库里没有」的题里，模型如实说没有的比例；",
        "- 首引命中率：引用的第一张卡片就是标准答案卡片的比例（比命中率严格）；",
        "- 引用精确率：引用的卡片里有多大比例是对的。专治「把检索结果全引上」刷分；",
        "- 引用命中率：引用的卡片里包含正确卡片的比例（口径最松，容易被 top-k 全引刷高）；",
        "- 虚构率：(该拒答却作答 + 答错且没引用到正确卡片) ÷ 总题数；",
        "- 覆盖率：没有被判为拒答的题占比。拒答越多，覆盖率越低——这是权衡，不是越高越好。",
        "",
        "> ⚠️ 以上全部是**脚本自动初判**，靠关键词和标记词匹配，会误判。",
        "> 论文里出现任何数字之前，必须逐题人工复核一遍，并在文中写明「自动初判 + 人工复核」。",
        "",
    ]
    if method == "b3":
        lines += [
            "> ⚠️ **B3 的数会虚高**：它直接把原文摘录整段贴出来，关键词必然命中，",
            "> 所以它的「事实正确率」不代表它答得好，只代表它抄得全。",
            "> 论文里 B3 的作用是证明「只贴原文不够，需要生成」，不能写成性能最好。",
            "",
        ]

    lines += [
        "## 逐题结果",
        "",
        "| 题号 | 类型 | 问题 | 标准答案 | 系统回答（截断） | 引用 | 事实 | 拒答 | 首引 | 精确率 |",
        "| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |",
    ]
    for r in rows:
        ans = r["answer"].replace("\n", " ").replace("|", "/")
        if len(ans) > 60:
            ans = ans[:60] + "…"
        fact = "—" if r["fact_ok"] is None else ("✅" if r["fact_ok"] else "❌")
        first = "—" if r["first_hit"] is None else ("✅" if r["first_hit"] else "❌")
        prec = "—" if r["cite_precision"] is None else f"{r['cite_precision'] * 100:.0f}%"
        lines.append(
            f'| {r["id"]} | {r["type"]} | {r["question"]} | {r["gold_answer"]} | '
            f'{ans} | {",".join(r["cited"]) or "—"} | {fact} | {"拒" if r["refused"] else "答"} | '
            f'{first} | {prec} |'
        )

    (base.with_suffix(".md")).write_text("\n".join(lines), encoding="utf-8")
    return base


# ---------------------------------------------------------------- 主流程
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--method", required=True, choices=["b1", "b2", "b3"])
    ap.add_argument("--topk", type=int, default=3)
    ap.add_argument("--limit", type=int, default=0, help="只跑前 N 题，用于试跑")
    args = ap.parse_args()

    cards = load_cards()
    items = load_eval()
    if args.limit:
        items = items[: args.limit]

    retriever = Retriever(cards)

    cfg = None
    if args.method in ("b1", "b2"):
        cfg = get_api_cfg()
        if cfg is None:
            print("[ERROR] 没检测到 OPENAI_API_KEY，b1/b2 跑不了。")
            print("        今晚想先验证流程，请用：python scripts/run_baseline.py --method b3")
            print("        本脚本不会在没调用大模型的情况下伪造 b1/b2 的结果。")
            return 2

    rows = []
    for i, item in enumerate(items, 1):
        if args.method == "b3":
            out = run_b3(item, retriever, args.topk)
        elif args.method == "b1":
            out = run_b1(item, retriever, cfg)
        else:
            out = run_b2(item, retriever, cfg, args.topk)

        fact_ok, refused, cite_hit, first_hit, precision = judge(item, out["answer"], out["cited"])
        rows.append({
            "id": item["id"],
            "type": item["type"],
            "question": item["question"],
            "gold_answer": item["gold_answer"],
            "gold_cards": item.get("gold_cards") or [],
            "answer": out["answer"],
            "cited": out["cited"],
            "fact_ok": fact_ok,
            "refused": refused,
            "cite_hit": cite_hit,
            "first_hit": first_hit,
            "cite_precision": precision,
        })
        print(f"[{i}/{len(items)}] {item['id']} done")

    summary = score(items, rows)
    meta = {
        "method": args.method,
        "time": time.strftime("%Y-%m-%d %H:%M:%S"),
        "model": cfg["model"] if cfg else "（未调用大模型，检索直出）",
        "topk": args.topk,
        "n_items": len(items),
    }
    base = write_outputs(args.method, rows, summary, meta)

    print("\n=== 汇总 ===")
    for k, v in summary.items():
        print(f"{k}: {v}")
    print(f"\n结果文件：{base}.md / {base}.json")
    return 0


if __name__ == "__main__":
    sys.exit(main())
