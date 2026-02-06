"""
IM Creator - AI Layout Engine
Version: 8.1.0

Uses Claude AI to analyze data and recommend optimal layouts for each slide.
NOW WITH INTEGRATED USAGE TRACKING!
"""

import os
import json
from typing import Dict, Any, Optional
from datetime import datetime
import anthropic

from utils import build_data_preview, get_default_layout_recommendation
from usage_tracker import get_tracker  # NEW: Import usage tracker

# ============================================================================
# LEGACY USAGE TRACKING (KEPT FOR BACKWARD COMPATIBILITY)
# ============================================================================
usage_stats = {
    "total_input_tokens": 0,
    "total_output_tokens": 0,
    "total_calls": 0,
    "total_cost_usd": 0.0,
    "session_start": datetime.now().isoformat(),
    "calls": []
}

PRICING = {
    "claude-sonnet-4-20250514": {"input": 0.003, "output": 0.015},
    "claude-3-5-sonnet-20241022": {"input": 0.003, "output": 0.015},
    "claude-3-opus-20240229": {"input": 0.015, "output": 0.075},
    "claude-3-haiku-20240307": {"input": 0.00025, "output": 0.00125}
}


def track_usage(model: str, input_tokens: int, output_tokens: int, purpose: str) -> float:
    """Track API usage and costs (LEGACY - kept for compatibility)"""
    pricing = PRICING.get(model, PRICING["claude-3-haiku-20240307"])
    cost_usd = (input_tokens / 1000 * pricing["input"]) + (output_tokens / 1000 * pricing["output"])
    
    usage_stats["total_input_tokens"] += input_tokens
    usage_stats["total_output_tokens"] += output_tokens
    usage_stats["total_calls"] += 1
    usage_stats["total_cost_usd"] += cost_usd
    
    usage_stats["calls"].append({
        "timestamp": datetime.now().isoformat(),
        "model": model,
        "input_tokens": input_tokens,
        "output_tokens": output_tokens,
        "cost_usd": f"{cost_usd:.6f}",
        "purpose": purpose
    })
    
    # Keep only last 100 calls
    if len(usage_stats["calls"]) > 100:
        usage_stats["calls"] = usage_stats["calls"][-100:]
    
    return cost_usd


def get_usage_stats() -> dict:
    """Get current usage statistics (LEGACY)"""
    return {
        **usage_stats,
        "total_cost_usd": f"{usage_stats['total_cost_usd']:.4f}"
    }


def reset_usage_stats():
    """Reset usage statistics (LEGACY)"""
    global usage_stats
    usage_stats = {
        "total_input_tokens": 0,
        "total_output_tokens": 0,
        "total_calls": 0,
        "total_cost_usd": 0.0,
        "session_start": datetime.now().isoformat(),
        "calls": []
    }


# ============================================================================
# AI LAYOUT ENGINE
# ============================================================================

async def analyze_data_for_layout(data: dict, slide_type: str) -> dict:
    """
    Analyze data and return AI-powered layout recommendations.
    
    Args:
        data: Form data dictionary
        slide_type: Type of slide to analyze
        
    Returns:
        Layout recommendation dictionary
    """
    # Build data preview
    try:
        data_preview = build_data_preview(data, slide_type)
    except Exception as e:
        print(f"Error building data preview for {slide_type}: {e}")
        data_preview = {"service_count": 0, "client_count": 0, "highlight_count": 0}
    
    # Check if AI is available
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key or os.environ.get("DISABLE_AI_LAYOUT") == "true":
        return get_default_layout_recommendation(slide_type, data_preview)
    
    prompt = f"""You are an expert presentation designer. Analyze this data for a "{slide_type}" slide.

DATA SUMMARY:
{json.dumps(data_preview, indent=2)}

Recommend the optimal design. Return ONLY valid JSON (no markdown):
{{
  "chart_type": "bar|pie|donut|progress|timeline|stacked-bar|none",
  "layout": "full-width|two-column|two-column-wide-left|two-column-wide-right|grid-2x2|grid-2x3",
  "font_adjustment": 0,
  "content_density": "low|medium|high",
  "primary_emphasis": "metrics|chart|text|mixed",
  "recommendations": ["suggestion1", "suggestion2"]
}}

Guidelines:
- Use pie/donut for 2-5 composition items
- Use bar for time series/revenue growth  
- Use progress bars for percentages/margins
- Use timeline for milestones/roadmap
- Use stacked-bar for revenue breakdown
- Use two-column for balanced content
- Use full-width for case studies or text-heavy content
- font_adjustment: 0 for normal, -1 for dense content, -2 for very dense
- Prioritize readability (12pt body minimum)"""

    try:
        client = anthropic.Anthropic(api_key=api_key)
        
        response = client.messages.create(
            model="claude-3-haiku-20240307",
            max_tokens=400,
            messages=[{"role": "user", "content": prompt}]
        )
        
        # v8.1.0: Single tracking via new usage tracker (removed legacy dual tracking)
        tracker = get_tracker()
        tracker.track_call(
            model="claude-3-haiku-20240307",
            input_tokens=response.usage.input_tokens,
            output_tokens=response.usage.output_tokens,
            purpose=f"AI Layout: {slide_type}"
        )
        
        text = response.content[0].text
        
        # Extract JSON from response
        import re
        json_match = re.search(r'\{[\s\S]*?\}', text)
        if json_match:
            parsed = json.loads(json_match.group())
            print(f"AI Layout for {slide_type}: {parsed}")
            return parsed
            
    except Exception as e:
        print(f"AI Layout fallback for {slide_type}: {e}")
    
    # Fallback to defaults
    return get_default_layout_recommendation(slide_type, data_preview)


def analyze_data_for_layout_sync(data: dict, slide_type: str) -> dict:
    """
    Synchronous version of analyze_data_for_layout.
    Used when async is not available.
    
    NOW WITH INTEGRATED USAGE TRACKING!
    """
    # Build data preview
    try:
        data_preview = build_data_preview(data, slide_type)
    except Exception as e:
        print(f"Error building data preview for {slide_type}: {e}")
        data_preview = {"service_count": 0, "client_count": 0, "highlight_count": 0}
    
    # Check if AI is available
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key or os.environ.get("DISABLE_AI_LAYOUT") == "true":
        return get_default_layout_recommendation(slide_type, data_preview)
    
    prompt = f"""You are an expert presentation designer. Analyze this data for a "{slide_type}" slide.

DATA SUMMARY:
{json.dumps(data_preview, indent=2)}

Recommend the optimal design. Return ONLY valid JSON (no markdown):
{{
  "chart_type": "bar|pie|donut|progress|timeline|stacked-bar|none",
  "layout": "full-width|two-column|two-column-wide-left|two-column-wide-right|grid-2x2|grid-2x3",
  "font_adjustment": 0,
  "content_density": "low|medium|high",
  "primary_emphasis": "metrics|chart|text|mixed",
  "recommendations": ["suggestion1", "suggestion2"]
}}

Guidelines:
- Use pie/donut for 2-5 composition items
- Use bar for time series/revenue growth  
- Use progress bars for percentages/margins
- Use timeline for milestones/roadmap
- Use stacked-bar for revenue breakdown
- Use two-column for balanced content
- Use full-width for case studies or text-heavy content
- font_adjustment: 0 for normal, -1 for dense content, -2 for very dense
- Prioritize readability (12pt body minimum)"""

    try:
        client = anthropic.Anthropic(api_key=api_key)
        
        response = client.messages.create(
            model="claude-3-haiku-20240307",
            max_tokens=400,
            messages=[{"role": "user", "content": prompt}]
        )
        
        # v8.1.0: Single tracking via new usage tracker (removed legacy dual tracking)
        tracker = get_tracker()
        tracker.track_call(
            model="claude-3-haiku-20240307",
            input_tokens=response.usage.input_tokens,
            output_tokens=response.usage.output_tokens,
            purpose=f"analyze_layout_{slide_type}"
        )
        
        text = response.content[0].text
        
        # Extract JSON from response
        import re
        json_match = re.search(r'\{[\s\S]*?\}', text)
        if json_match:
            parsed = json.loads(json_match.group())
            print(f"AI Layout for {slide_type}: {parsed}")
            return parsed
            
    except Exception as e:
        print(f"AI Layout fallback for {slide_type}: {e}")
    
    # Fallback to defaults
    return get_default_layout_recommendation(slide_type, data_preview)