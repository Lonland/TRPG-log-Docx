from difflib import SequenceMatcher
from datetime import datetime

def content_to_text(content_lines):
    return "\n".join(content_lines).strip()

def is_same_hour(t1, t2):
    return t1.strftime("%Y-%m-%d %H") == t2.strftime("%Y-%m-%d %H")

def similarity(a, b):
    return SequenceMatcher(None, a, b).ratio()

def deduplicate_messages(messages):
    # 首先将所有时间字段转换为 datetime 对象（临时加入，不改变原结构）
    for msg in messages:
        msg["_dt"] = datetime.strptime(msg["time"], "%Y-%m-%d %H:%M:%S")

    # 按时间排序（越早的排前）
    messages.sort(key=lambda m: m["_dt"])

    result = []
    seen = []

    for i, msg in enumerate(messages):
        name_i = msg["name"]
        time_i = msg["_dt"]
        text_i = content_to_text(msg["content"])

        is_duplicate = False
        for j in reversed(range(len(result))):  # 从后往前比较已有保留的
            prev = result[j]
            if prev["name"] != name_i:
                continue
            if not is_same_hour(time_i, prev["_dt"]):
                continue

            text_j = content_to_text(prev["content"])
            sim = similarity(text_i, text_j)
            if sim >= 0.9:
                # 当前是重复的，标记为丢弃
                is_duplicate = True
                break

        if not is_duplicate:
            result.append(msg)

    # 去掉临时字段
    for msg in result:
        msg.pop("_dt", None)

    return result
