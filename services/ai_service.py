import json
import re
from typing import Any, Dict, List

import requests
from fastapi import HTTPException


def extract_json_from_text(text: str) -> Dict[str, Any]:
    text = text.strip()

    # 1. Strip markdown code block fences (```json ... ``` or ``` ... ```) first before parsing
    fence_match = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", text, re.IGNORECASE)
    cleaned_text = fence_match.group(1).strip() if fence_match else text

    # Try parsing stripped text
    try:
        data = json.loads(cleaned_text)
        if isinstance(data, dict):
            return data
    except json.JSONDecodeError:
        pass

    # If stripped text didn't parse, try original text
    if cleaned_text != text:
        try:
            data = json.loads(text)
            if isinstance(data, dict):
                return data
        except json.JSONDecodeError:
            pass

    # 2. If json decode fails, search for the outer { ... }
    match = re.search(r"\{[\s\S]*\}", cleaned_text) or re.search(r"\{[\s\S]*\}", text)
    if not match:
        raise ValueError("AI 返回内容中未找到 JSON")

    try:
        data = json.loads(match.group(0))
    except json.JSONDecodeError as exc:
        raise ValueError(f"解析 JSON 失败: {str(exc)}")

    if not isinstance(data, dict):
        raise ValueError("AI 返回的 JSON 必须是字典对象")

    return data


def call_ai_generate_json(
    prompt: str,
    api_url: str,
    api_key: str,
    model: str,
    variables: List[str],
    stream: bool = False,
    temperature: float = 0.2,
) -> Dict[str, str]:
    schema = {k: "" for k in variables}

    system_prompt = (
        "你是一个严谨的 JSON 生成器。"
        "只允许输出一个标准 JSON 对象，不允许任何额外文字、Markdown 或代码块。"
        "必须包含给定字段，且字段名完全一致。"
    )

    user_prompt = (
        f"用户提示词：{prompt}\n\n"
        f"请基于提示词生成 JSON，字段必须严格如下：\n"
        f"{json.dumps(schema, ensure_ascii=False, indent=2)}\n\n"
        f"要求：\n"
        f"1) 仅输出 JSON 对象\n"
        f"2) 所有字段必须存在\n"
        f"3) 字段值使用中文详细填充对应内容"
    )

    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": temperature,
        "response_format": {"type": "json_object"},
    }
    if stream:
        payload["stream"] = True

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    try:
        resp = requests.post(api_url, headers=headers, json=payload, timeout=90, stream=stream)
    except requests.RequestException as exc:
        raise HTTPException(status_code=502, detail=f"AI 请求失败: {str(exc)}")

    if resp.status_code >= 400:
        error_detail = ""
        try:
            err_json = resp.json()
            if isinstance(err_json, dict):
                error_obj = err_json.get("error", err_json)
                if isinstance(error_obj, dict):
                    error_detail = error_obj.get("message") or str(error_obj)
                else:
                    error_detail = str(error_obj)
        except Exception:
            pass
        if not error_detail:
            error_detail = resp.text[:1000] if hasattr(resp, "text") else ""

        if resp.status_code == 400:
            raise HTTPException(
                status_code=400,
                detail=f"AI API 请求参数错误 (400)：{error_detail}",
            )
        elif resp.status_code == 401:
            raise HTTPException(
                status_code=401,
                detail=f"AI API 认证失败 (401)，请检查 API Key 是否正确：{error_detail}",
            )
        elif resp.status_code == 403:
            raise HTTPException(
                status_code=403,
                detail=f"AI API 访问受限 (403)，请检查账户权限或余额：{error_detail}",
            )
        elif resp.status_code == 404:
            raise HTTPException(
                status_code=404,
                detail=f"AI API 端点或模型不存在 (404)，请检查 API URL 及模型名称：{error_detail}",
            )
        elif resp.status_code == 429:
            raise HTTPException(
                status_code=429,
                detail=f"AI API 请求频次受限或额度不足 (429)：{error_detail}",
            )
        else:
            raise HTTPException(
                status_code=resp.status_code,
                detail=f"AI 接口错误 ({resp.status_code})：{error_detail}",
            )

    content_type = resp.headers.get("content-type", "").lower()
    is_event_stream = "text/event-stream" in content_type

    if stream or is_event_stream:
        content_chunks = []
        for raw_line in resp.iter_lines():
            if not raw_line:
                continue
            if isinstance(raw_line, bytes):
                line = raw_line.decode("utf-8", errors="replace").strip()
            else:
                line = str(raw_line).strip()
            if not line:
                continue
            if line.startswith("data:"):
                data_str = line[len("data:"):].strip()
                if data_str == "[DONE]":
                    break
                try:
                    chunk_json = json.loads(data_str)
                    choices = chunk_json.get("choices", [])
                    if choices:
                        delta = choices[0].get("delta", {})
                        content_piece = delta.get("content")
                        if content_piece:
                            content_chunks.append(content_piece)
                except json.JSONDecodeError:
                    continue
        content = "".join(content_chunks)
        if not content:
            raise HTTPException(status_code=500, detail="AI 流式响应未获取到有效内容")
    else:
        try:
            data = resp.json()
            content = data["choices"][0]["message"]["content"]
        except Exception:
            raise HTTPException(status_code=500, detail=f"AI 响应格式异常: {resp.text[:1000]}")

    if isinstance(content, dict):
        parsed = content
    elif content is None:
        raise HTTPException(status_code=500, detail="AI 响应内容为空")
    else:
        try:
            parsed = extract_json_from_text(str(content))
        except Exception as exc:
            raise HTTPException(
                status_code=500,
                detail=f"AI 未返回合法 JSON: {str(exc)}; 原始内容: {str(content)[:800]}",
            )

    missing = [k for k in variables if k not in parsed]
    if missing:
        raise HTTPException(status_code=500, detail=f"AI 返回缺少字段: {missing}")

    for key in variables:
        if parsed.get(key) is None:
            parsed[key] = ""
        else:
            parsed[key] = str(parsed[key])

    return parsed
