"""
Gemini Multimodal Client with Resilient Fallback Routing.
Calls the official Google Generative Language API via direct HTTPS REST calls,
supporting structured JSON schemas, text extraction, and model failover.
"""

import os
import json
import time
import urllib.request
import urllib.error
from typing import Dict, Any, Optional, List

# Load local .env if present
env_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), '.env')
if os.path.exists(env_path):
    try:
        with open(env_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    k, v = line.split('=', 1)
                    k, v = k.strip(), v.strip().strip('"\'')
                    if k not in os.environ:
                        os.environ[k] = v
    except Exception:
        pass

# Preferred models in priority order
CANDIDATE_MODELS = [
    "gemini-3.5-flash",
    "gemini-3.5-flash-lite",
    "gemini-3-flash-preview",
    "gemini-3.8-flash",
]


def get_api_key() -> str:
    """Retrieves API key from Streamlit session, secrets, or environment."""
    try:
        import streamlit as st
        if hasattr(st, "session_state") and st.session_state.get("gemini_api_key"):
            return st.session_state["gemini_api_key"].strip()
        if hasattr(st, "secrets") and "GEMINI_API_KEY" in st.secrets:
            return str(st.secrets["GEMINI_API_KEY"]).strip()
    except Exception:
        pass
    return os.environ.get("GEMINI_API_KEY", "").strip()


def call_gemini(
    prompt: str,
    system_instruction: Optional[str] = None,
    response_schema: Optional[Dict[str, Any]] = None,
    max_retries: int = 3,
    timeout: int = 45
) -> str:
    """
    Executes a prompt against Gemini with automatic model failover and retries.
    Returns raw text or JSON string from candidate model.
    """
    api_key = get_api_key()
    
    payload: Dict[str, Any] = {
        "contents": [{"parts": [{"text": prompt}]}]
    }
    
    if system_instruction:
        payload["systemInstruction"] = {
            "parts": [{"text": system_instruction}]
        }
        
    gen_config: Dict[str, Any] = {
        "temperature": 0.1,
        "maxOutputTokens": 8192
    }
    
    if response_schema:
        gen_config["response_mime_type"] = "application/json"
        gen_config["response_schema"] = response_schema
        
    payload["generationConfig"] = gen_config
    data_bytes = json.dumps(payload).encode("utf-8")
    
    last_error = None
    for model in CANDIDATE_MODELS:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
        
        for attempt in range(max_retries):
            try:
                req = urllib.request.Request(
                    url,
                    data=data_bytes,
                    headers={"Content-Type": "application/json"}
                )
                with urllib.request.urlopen(req, timeout=timeout) as resp:
                    if resp.status == 200:
                        res_json = json.loads(resp.read().decode("utf-8"))
                        candidates = res_json.get("candidates", [])
                        if candidates and "content" in candidates[0]:
                            parts = candidates[0]["content"].get("parts", [])
                            if parts and "text" in parts[0]:
                                return parts[0]["text"]
            except urllib.error.HTTPError as e:
                err_body = e.read().decode("utf-8", errors="ignore")
                last_error = f"{model} HTTP {e.code}: {err_body[:200]}"
                # If rate limited (429) or temporary server error (503/500), sleep and retry or fall back
                if e.code in [429, 503, 500]:
                    time.sleep(1.5 * (attempt + 1))
                    continue
                elif e.code == 404:
                    # Model not found or deprecated, try next model in list
                    break
            except Exception as e:
                last_error = f"{model} Error: {str(e)}"
                time.sleep(1.0)
                continue
                
    raise RuntimeError(f"All Gemini models failed. Last error: {last_error}")


def call_gemini_json(
    prompt: str,
    response_schema: Dict[str, Any],
    system_instruction: Optional[str] = None
) -> Dict[str, Any]:
    """Generates structured JSON directly matching the provided schema."""
    raw_text = call_gemini(
        prompt=prompt,
        system_instruction=system_instruction,
        response_schema=response_schema
    )
    # Parse JSON
    try:
        return json.loads(raw_text)
    except json.JSONDecodeError:
        # Strip potential markdown fences if returned
        clean = raw_text.strip()
        if clean.startswith("```json"):
            clean = clean[7:]
        if clean.startswith("```"):
            clean = clean[3:]
        if clean.endswith("```"):
            clean = clean[:-3]
        return json.loads(clean.strip())
