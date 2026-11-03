import os
import re
import json
import time
import threading
import pandas as pd
from openai import OpenAI
import httpx
from concurrent.futures import ThreadPoolExecutor, as_completed

# 直连内网部署服务：清掉环境代理并排除内网地址，否则内网请求会被 HIS 代理拦截返回 HTML 页面
for _proxy_key in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY",
                   "http_proxy", "https_proxy", "all_proxy"):
    os.environ.pop(_proxy_key, None)
os.environ["NO_PROXY"] = ",".join(["7.216.57.45", "127.0.0.1", "localhost"])

# ==================== 输入输出配置 ====================
INPUT_FILE = r"C:\Users\h00993790\Desktop\thinker_英语忠实度结果.xlsx"
SHEET_NAME = ""      # 留空=取第一个工作表
USER_COL = "Input"          # 用户输入/参照列；分组模式(GROUP_MODE=True)下改为按此列(例如 Question)分组
TARGET_COL = "Scatter"      # 待评转述列
GROUP_MODE = True           # False=每行一个 TARGET_COL 单独评分；True=按 USER_COL 分组，同组 TARGET_COL 用 JOIN_SEP 连成一句后评分
JOIN_SEP = " "              # 分组模式拼接分隔符（空格/顿号等）
STRIP_BOV = True             # 剥离 <bov>/<eov> 标签后再送评，避免干扰
OUTPUT_FILE = r"C:\Users\h00993790\Desktop\thinker_英语综合评分结果.xlsx"  # 独立结果表（不写回原表）
NUM_LIMIT = None             # None=全部；填数字=只处理前 N 行（分组模式下为前 N 组）

# API 配置（本机部署的 deepseek-v41，直连不走代理）
API_KEY = "148db3a5fc22f0f1d1b2f9a2fe8dc0c3"
BASE_URL = "http://7.216.57.45:50424/v1"
MODEL = "deepseek-v41"
ENABLE_THINKING = False   # 评分任务不需要推理，关闭以加速；如需可改 True

MAX_RETRIES = 3
MAX_WORKERS = 8    # 并发线程数；若接口限流/卡住可调回 2-3
REQUEST_TIMEOUT = 120

RESET = True           # True=从头重新打分（清空旧分数+不用旧进度）；打分完成后可改 False 复用新进度
PROGRESS_FILE = "mos评分_progress_deepseek.json"

# ==================== MOS 评分标准（完全按用户提供的标准 JSON 定义） ====================
DIMENSIONS = [
    {
        "id": "intent_execution",
        "name": "意图理解和执行",
        "weight": 0.2,
        "min_score": 1,
        "max_score": 5,
        "system_prompt": "你是严格的对话系统评测员。只评价给定用户输入与输出内容，不补充输出中不存在的信息。",
        "prompt": (
            "请评价输出内容是否准确回应并执行了用户输入中的意图和要求。\n\n"
            "语言：{language}\n"
            "意图维度：{dimensions}\n"
            "用户输入：\n{user_input}\n\n"
            "输出内容：\n{acm_output}\n\n"
            "评分规则：\n"
            "5=完整捕获并执行所有明确指令，准确处理关键对象、时间、数量和约束。\n"
            "4=遵循了大部分指令，仅有轻微遗漏或小错误，核心任务仍完成。\n"
            "3=只执行部分指令，存在明显遗漏，但对核心意图有有效响应。\n"
            "2=遗漏或错误处理主要步骤，核心任务基本未完成。\n"
            "1=没有回应核心意图、答非所问、拒绝有效请求，或没有有效输出。\n\n"
            "只输出 JSON：{{\"reason\":\"不超过80字的评分依据\",\"score\":1到5的整数}}"
        ),
    },
    {
        "id": "topic_relevance",
        "name": "话题相关",
        "weight": 0.2,
        "min_score": 1,
        "max_score": 5,
        "system_prompt": "你是严格的对话系统评测员。相关性评分必须以给定的用户输入和输出内容为依据。",
        "prompt": (
            "请根据用户输入，评价输出内容与用户话题及核心诉求的相关性。\n\n"
            "语言：{language}\n"
            "用户输入：\n{user_input}\n\n"
            "输出内容：\n{acm_output}\n\n"
            "评分规则：\n"
            "5=直接回应核心诉求，信息有价值且没有偏题或冗余。\n"
            "4=主体相关，细节回应稍弱或有少量通用表达，但没有偏离主题。\n"
            "3=只回应部分核心内容，其余较泛化，或引入弱相关内容。\n"
            "2=主要复述关键词，实质回应很少，存在明显话题转移或大量模板话术。\n"
            "1=内容与用户话题、情绪或意图完全不匹配，或严重答非所问。\n\n"
            "只输出 JSON：{{\"reason\":\"不超过80字的评分依据\",\"score\":1到5的整数}}"
        ),
    },
    {
        "id": "fluency",
        "name": "语句通顺",
        "weight": 0.2,
        "min_score": 1,
        "max_score": 5,
        "system_prompt": "你是严格的多语言表达质量评测员。请按当前回答语言评价语法、自然度和可理解性。",
        "prompt": (
            "请评价输出内容是否语句通顺、自然且易于理解。\n\n"
            "语言：{language}\n"
            "用户输入：\n{user_input}\n\n"
            "输出内容：\n{acm_output}\n\n"
            "评分规则：\n"
            "5=语法正确，表达自然流畅，标点和语气恰当，符合该语言日常表达。\n"
            "4=基本通顺，偶有轻微语法或措辞瑕疵，但不影响理解。\n"
            "3=整体可读，存在少量错误或生硬表达，需要稍作理解。\n"
            "2=有明显语法、语序或用词问题，理解较费力，机械感强。\n"
            "1=错误密集或严重不通顺，难以理解核心意思。\n\n"
            "只输出 JSON：{{\"reason\":\"不超过80字的评分依据\",\"score\":1到5的整数}}"
        ),
    },
    {
        "id": "consistency",
        "name": "前后一致",
        "weight": 0.2,
        "min_score": 1,
        "max_score": 5,
        "system_prompt": "你是严格的对话一致性评测员。当前是单轮评测，只评价回答内部以及回答与用户输入之间的一致性。",
        "prompt": (
            "请结合用户输入，评价输出内容的逻辑一致性。不要假设未提供的历史对话。\n\n"
            "用户输入：\n{user_input}\n\n"
            "输出内容：\n{acm_output}\n\n"
            "评分规则：\n"
            "5=与用户输入衔接自然，回答内部逻辑严密，事实、对象和指代前后一致。\n"
            "4=整体一致，偶有细节呼应不足，但没有实质矛盾。\n"
            "3=基本连贯，存在轻微逻辑跳跃、指代不清或细节不一致。\n"
            "2=衔接生硬，存在明显逻辑断裂、对象混淆或自相矛盾。\n"
            "1=与输入完全脱节或回答内部严重矛盾，无法形成有效对话。\n\n"
            "只输出 JSON：{{\"reason\":\"不超过80字的评分依据\",\"score\":1到5的整数}}"
        ),
    },
    {
        "id": "conciseness",
        "name": "简明扼要",
        "weight": 0.2,
        "min_score": 1,
        "max_score": 5,
        "system_prompt": "你是严格的对话简洁性评测员。请结合任务复杂度判断长度是否合适，避免把必要步骤误判为冗余。",
        "prompt": (
            "请结合用户输入，评价输出内容是否简明且信息密度适当。\n\n"
            "用户输入：\n{user_input}\n\n"
            "输出内容：\n{acm_output}\n\n"
            "评分规则：\n"
            "5=长度与任务匹配，简洁自然，无冗余重复，同时保留必要信息。\n"
            "4=较为简洁，仅有轻微冗余或少量可压缩表达。\n"
            "3=基本可接受，但有重复、偏长、偏短或略显机械的问题。\n"
            "2=明显冗长、重复、过度格式化，或过度简略而影响信息完整性。\n"
            "1=大量无关内容、严重重复，或短到无法形成有效回答。\n\n"
            "只输出 JSON：{{\"reason\":\"不超过80字的评分依据\",\"score\":1到5的整数}}"
        ),
    },
]

DIM_KEYS = [d["id"] for d in DIMENSIONS]
COL_NAMES = ["意图执行", "话题相关", "语句通顺", "前后一致", "简明扼要", "MOS总分"]


def detect_language(text: str) -> str:
    if not text:
        return "中文"
    ar = len(re.findall(r"[\u0600-\u06FF]", text))
    zh = len(re.findall(r"[\u4e00-\u9fff]", text))
    if ar > 0 and ar >= zh:
        return "阿拉伯语"
    if zh > 0:
        return "中文"
    return "英文"


def fill_template(template: str, user_input: str, acm_output: str, dim_name: str, language: str) -> str:
    text = template.replace("{{", "{").replace("}}", "}")
    return (
        text.replace("{language}", language)
        .replace("{dimensions}", dim_name)
        .replace("{user_input}", user_input)
        .replace("{acm_output}", acm_output)
    )


def build_user_prompt(user_input: str, acm_output: str):
    language = detect_language(user_input + " " + acm_output)
    parts = []
    for d in DIMENSIONS:
        parts.append(
            f"维度：{d['name']}（id={d['id']}，权重={d['weight']}）\n"
            f"评测人设定：{d['system_prompt']}\n"
            f"语言：{language}\n"
            f"评测指令：\n{fill_template(d['prompt'], user_input, acm_output, d['name'], language)}"
        )
    return (
        "请对以下用户输入和输出内容，按下列五个维度逐个评分。\n\n"
        f"用户输入：\n{user_input}\n\n输出内容：\n{acm_output}\n\n"
        "评测内容语言：" + language + "\n\n"
        + "\n\n=====\n\n".join(parts)
        + "\n\n请最终只输出一个 JSON 对象，键为五个维度 id，值为各维度要求的 {\"reason\":...,\"score\":...}，不要输出其他内容。"
    )


def get_client():
    return OpenAI(
        api_key=API_KEY,
        base_url=BASE_URL,
        http_client=httpx.Client(verify=False, trust_env=False, timeout=httpx.Timeout(REQUEST_TIMEOUT, connect=10.0)),
        timeout=REQUEST_TIMEOUT,
    )


def score_chunk(user_input: str, acm_output: str, client) -> dict:
    params = {
        "model": MODEL,
        "messages": [
            {"role": "user", "content": build_user_prompt(user_input, acm_output)},
        ],
        "temperature": 0.0,
    }

    if ENABLE_THINKING:
        params["extra_body"] = {"thinking": {"type": "enabled"}}
        params["stream"] = True
        completion = client.chat.completions.create(**params)
        full_content = ""
        for chunk in completion:
            if not chunk.choices:
                continue
            delta = chunk.choices[0].delta
            if hasattr(delta, "content") and delta.content:
                full_content += delta.content
        content_text = full_content.strip()
    else:
        params["extra_body"] = {"thinking": {"type": "disabled"}}
        completion = client.chat.completions.create(**params)
        content_text = completion.choices[0].message.content.strip()

    try:
        return json.loads(content_text)
    except (json.JSONDecodeError, ValueError):
        start = content_text.find("{")
        end = content_text.rfind("}")
        if start != -1 and end > start:
            try:
                return json.loads(content_text[start:end + 1])
            except (json.JSONDecodeError, ValueError):
                pass
    raise ValueError(f"模型输出无法解析为 JSON: {content_text[:200]}")


def parse_result(result: dict):
    scores = []
    for key in DIM_KEYS:
        item = result.get(key)
        if not isinstance(item, dict) or "score" not in item:
            raise ValueError(f"结果缺少维度 {key}: {str(result)[:200]}")
        try:
            s = int(item["score"])
        except (TypeError, ValueError):
            raise ValueError(f"维度 {key} 分数非整数: {item}")
        if not (1 <= s <= 5):
            raise ValueError(f"维度 {key} 分数越界: {s}")
        scores.append(s)
    mos = round(sum(s * d["weight"] for s, d in zip(scores, DIMENSIONS)), 1)
    return scores, mos


def clean_text(value: str) -> str:
    if STRIP_BOV:
        value = re.sub(r"</?(?:bov|eov)>", "", value, flags=re.I)
    return value.strip()


def load_tasks(df: pd.DataFrame):
    """返回 [(task_id, user_input, target_text), ...]。逐行模式按原索引；分组模式按 USER_COL 保持首次出现顺序拼接。"""
    tasks = []
    if GROUP_MODE:
        grouped = {}
        order = []
        for idx, row in df.iterrows():
            key = str(row.get(USER_COL, "") or "").strip()
            tgt = clean_text(str(row.get(TARGET_COL, "") or ""))
            if not tgt:
                continue
            if key not in grouped:
                grouped[key] = []
                order.append(key)
            grouped[key].append(tgt)
        for i, key in enumerate(order):
            tasks.append((f"g{i}", key, JOIN_SEP.join(grouped[key])))
    else:
        for idx, row in df.iterrows():
            user = clean_text(str(row.get(USER_COL, "") or ""))
            target = clean_text(str(row.get(TARGET_COL, "") or ""))
            if not target:
                continue
            tasks.append((str(idx), user, target))
    return tasks


def print_mos_distribution(df_scores):
    vals = pd.to_numeric(df_scores["MOS总分"], errors="coerce").dropna()
    if len(vals) == 0:
        print("📊 暂无可统计的 MOS 数据")
        return
    total = len(vals)
    print("📊 MOS 分布统计")
    print(f"   有效样本: {total} 条")
    print(f"   min={vals.min():.2f}  max={vals.max():.2f}  mean={vals.mean():.2f}  std={vals.std():.2f}")
    scale = max(1, total // 50)
    buckets = [(0, 3, "<3"), (3, 3.5, "3~3.4"), (3.5, 4, "3.5~3.9"), (4, 4.5, "4~4.4"), (4.5, 5.01, "4.5~5")]
    for lo, hi, label in buckets:
        if hi > 5:
            cnt = int(((vals >= lo) & (vals <= 5)).sum())
        else:
            cnt = int(((vals >= lo) & (vals < hi)).sum())
        bar = "#" * (cnt // scale)
        print(f"   MOS {label:<7}: {cnt:5d}  {cnt / total * 100:5.1f}%  {bar}")
    print("   --- 各维度平均分 ---")
    for col in COL_NAMES[:-1]:
        sub = pd.to_numeric(df_scores[col], errors="coerce").dropna()
        if len(sub):
            print(f"   {col}: mean={sub.mean():.2f}")


def main():
    print("=" * 70)
    mode = "按" + USER_COL + "分组拼接" if GROUP_MODE else "逐行"
    print(f"📊 MOS 打分工具（deepseek-v41，本机部署，{mode}模式）")
    print(f"📁 输入: {INPUT_FILE} (sheet={SHEET_NAME or '首个'}, {USER_COL}→{TARGET_COL})")
    print(f"💾 输出: {OUTPUT_FILE}（独立结果表，不写回原表）")
    print(f"🚀 并发: {MAX_WORKERS} 线程")
    print("=" * 70)

    if not os.path.exists(INPUT_FILE):
        print(f"❌ 输入文件不存在: {INPUT_FILE}")
        return

    df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME if SHEET_NAME else 0)
    if USER_COL not in df.columns or TARGET_COL not in df.columns:
        print(f"❌ 表格缺少 {USER_COL} 或 {TARGET_COL} 列，实际列: {list(df.columns)}")
        return
    print(f"📖 {USER_COL}/{TARGET_COL} 共 {len(df)} 行")

    tasks = load_tasks(df)
    if NUM_LIMIT is not None:
        tasks = tasks[:NUM_LIMIT]
    print(f"🔄 待评分任务: {len(tasks)} 条")

    # 初始化结果容器
    if GROUP_MODE:
        result_rows = {tid: {USER_COL: user, f"{TARGET_COL}_拼接": target}
                       for tid, user, target in tasks}
    else:
        for col in COL_NAMES:
            if col not in df.columns:
                df[col] = pd.Series([None] * len(df), dtype=object)
            else:
                df[col] = df[col].astype(object)

    progress = {}
    if not RESET and os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                progress = json.load(f)
        except Exception:
            progress = {}

    todo = [(tid, user, target) for tid, user, target in tasks if tid not in progress]
    print(f"🔄 待评: {len(todo)} 条（进度 {len(progress)} 条）")

    if todo:
        def worker(task):
            tid, user, target = task
            client = get_client()
            last_err = None
            for attempt in range(1, MAX_RETRIES + 1):
                try:
                    result = score_chunk(user, target, client)
                    scores, mos = parse_result(result)
                    return tid, scores, mos, None, target
                except Exception as e:
                    last_err = e
                    time.sleep(2 ** (attempt - 1))
            return tid, None, None, str(last_err), target

        lock = threading.Lock()
        done_count = 0
        failed = 0

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
            futures = {ex.submit(worker, t): t for t in todo}
            for fut in as_completed(futures):
                tid, scores, mos, err, target = fut.result()
                with lock:
                    if err is None:
                        if GROUP_MODE:
                            result_rows[tid].update({c: s for c, s in zip(COL_NAMES[:-1], scores)})
                            result_rows[tid][COL_NAMES[-1]] = mos
                        else:
                            idx = int(tid)
                            for col, s in zip(COL_NAMES[:-1], scores):
                                df.at[idx, col] = s
                            df.at[idx, COL_NAMES[-1]] = mos
                        progress[tid] = {"scores": scores, "mos": mos}
                        status = "✅"
                        detail = f"MOS={mos} 明细={scores}"
                    else:
                        failed += 1
                        progress[tid] = {"error": err}
                        status = "❌"
                        detail = f"错误: {(err or '')[:300]}"
                    done_count += 1
                    print(f"   {status} [{done_count}/{len(todo)}] {tid} {detail}", flush=True)

                    if done_count % 50 == 0 or done_count == len(todo):
                        with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
                            json.dump(progress, f, ensure_ascii=False, indent=2)
                        # 增量写结果表
                        if GROUP_MODE:
                            res_df = pd.DataFrame.from_dict(result_rows, orient="index")
                        else:
                            res_df = df
                        with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
                            res_df.to_excel(writer, sheet_name="结果", index=False)
                        print(f"   💾 已保存结果表，进度 {done_count}/{len(todo)} (失败: {failed})", flush=True)

        with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
            json.dump(progress, f, ensure_ascii=False, indent=2)
        print(f"   ✅ 打分完成，失败 {failed} 条（失败条目下次运行自动重试）")
    else:
        print("   ✅ 全部已完成（复用进度）")

    # 最终写独立结果表
    if GROUP_MODE:
        final_df = pd.DataFrame.from_dict(result_rows, orient="index")
    else:
        final_df = df
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        final_df.to_excel(writer, sheet_name="结果", index=False)
    print(f"💾 结果表已写出: {OUTPUT_FILE}")
    print("=" * 70)
    print_mos_distribution(final_df)
    print("=" * 70)
    print("✅ 完成！")


if __name__ == "__main__":
    main()
