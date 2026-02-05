"""
IM Creator - Anthropic API Usage Tracker
Version: 8.0.0

Tracks API usage, token counts, and costs for Anthropic Claude API calls.
Integrates with frontend Usage button.
"""

from typing import Dict, List, Optional
from datetime import datetime
import json
from pathlib import Path

# ============================================================================
# ANTHROPIC API PRICING (per 1K tokens)
# ============================================================================

PRICING = {
    "claude-opus-4-5-20251101": {
        "input": 0.015,
        "output": 0.075,
        "name": "Claude Opus 4.5"
    },
    "claude-sonnet-4-5-20250929": {
        "input": 0.003,
        "output": 0.015,
        "name": "Claude Sonnet 4.5"
    },
    "claude-sonnet-4-20250514": {
        "input": 0.003,
        "output": 0.015,
        "name": "Claude Sonnet 4"
    },
    "claude-haiku-4-5-20251001": {
        "input": 0.00025,
        "output": 0.00125,
        "name": "Claude Haiku 4.5"
    },
    "claude-3-5-sonnet-20241022": {
        "input": 0.003,
        "output": 0.015,
        "name": "Claude 3.5 Sonnet"
    },
    "claude-3-opus-20240229": {
        "input": 0.015,
        "output": 0.075,
        "name": "Claude 3 Opus"
    },
    "claude-3-haiku-20240307": {
        "input": 0.00025,
        "output": 0.00125,
        "name": "Claude 3 Haiku"
    }
}


# ============================================================================
# USAGE TRACKER CLASS
# ============================================================================

class UsageTracker:
    """
    Tracks Anthropic API usage and costs
    
    Features:
    - Token counting (input + output)
    - Cost calculation per model
    - Session-based tracking
    - Detailed call logs
    - CSV export functionality
    - Reset functionality
    - Persistent storage
    """
    
    def __init__(self, storage_path: Optional[Path] = None):
        """
        Initialize usage tracker
        
        Args:
            storage_path: Path to store usage data (default: /tmp/anthropic_usage.json)
        """
        self.storage_path = storage_path or Path("/tmp/anthropic_usage.json")
        self.stats = self._load_stats()
    
    def _load_stats(self) -> Dict:
        """Load existing stats from disk or create new"""
        if self.storage_path.exists():
            try:
                with open(self.storage_path, 'r') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Warning: Could not load usage stats: {e}")
        
        # Return fresh stats
        return {
            "total_calls": 0,
            "total_input_tokens": 0,
            "total_output_tokens": 0,
            "total_cost_usd": 0.0,
            "session_start": datetime.now().isoformat(),
            "calls": [],
            "by_purpose": {},
            "by_model": {}
        }
    
    def _save_stats(self):
        """Persist stats to disk"""
        try:
            # Ensure directory exists
            self.storage_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(self.storage_path, 'w') as f:
                json.dump(self.stats, f, indent=2)
        except Exception as e:
            print(f"Warning: Could not save usage stats: {e}")
    
    def track_call(
        self,
        model: str,
        input_tokens: int,
        output_tokens: int,
        purpose: str = "general"
    ) -> Dict:
        """
        Track an API call and calculate cost
        
        Args:
            model: Model identifier (e.g., "claude-sonnet-4-20250514")
            input_tokens: Number of input tokens used
            output_tokens: Number of output tokens generated
            purpose: Purpose of the call (e.g., "analyze_layout", "generate_content")
        
        Returns:
            Dictionary with call details and cost
        """
        # Get pricing for this model (fallback to Haiku if unknown)
        pricing = PRICING.get(model, PRICING["claude-3-haiku-20240307"])
        
        # Calculate costs
        input_cost = (input_tokens / 1000) * pricing["input"]
        output_cost = (output_tokens / 1000) * pricing["output"]
        total_cost = input_cost + output_cost
        
        # Create call record
        call_record = {
            "timestamp": datetime.now().isoformat(),
            "model": model,
            "model_name": pricing.get("name", model),
            "purpose": purpose,
            "input_tokens": input_tokens,
            "output_tokens": output_tokens,
            "total_tokens": input_tokens + output_tokens,
            "cost_usd": round(total_cost, 6)
        }
        
        # Update global totals
        self.stats["total_calls"] += 1
        self.stats["total_input_tokens"] += input_tokens
        self.stats["total_output_tokens"] += output_tokens
        self.stats["total_cost_usd"] = round(self.stats["total_cost_usd"] + total_cost, 6)
        
        # Update by-purpose breakdown
        if purpose not in self.stats["by_purpose"]:
            self.stats["by_purpose"][purpose] = {
                "calls": 0,
                "tokens": 0,
                "cost": 0.0
            }
        
        self.stats["by_purpose"][purpose]["calls"] += 1
        self.stats["by_purpose"][purpose]["tokens"] += input_tokens + output_tokens
        self.stats["by_purpose"][purpose]["cost"] = round(
            self.stats["by_purpose"][purpose]["cost"] + total_cost, 6
        )
        
        # Update by-model breakdown
        model_name = pricing.get("name", model)
        if model_name not in self.stats["by_model"]:
            self.stats["by_model"][model_name] = {
                "calls": 0,
                "tokens": 0,
                "cost": 0.0
            }
        
        self.stats["by_model"][model_name]["calls"] += 1
        self.stats["by_model"][model_name]["tokens"] += input_tokens + output_tokens
        self.stats["by_model"][model_name]["cost"] = round(
            self.stats["by_model"][model_name]["cost"] + total_cost, 6
        )
        
        # Add to call log (keep last 1000 calls only)
        self.stats["calls"].append(call_record)
        if len(self.stats["calls"]) > 1000:
            self.stats["calls"] = self.stats["calls"][-1000:]
        
        # Save to disk
        self._save_stats()
        
        return call_record
    
    def get_stats(self) -> Dict:
        """
        Get current usage statistics
        
        Returns:
            Dictionary with all usage statistics including totals and averages
        """
        return {
            **self.stats,
            "session_duration_hours": self._get_session_duration(),
            "average_cost_per_call": round(
                self.stats["total_cost_usd"] / max(1, self.stats["total_calls"]), 6
            ),
            "average_tokens_per_call": round(
                (self.stats["total_input_tokens"] + self.stats["total_output_tokens"]) / 
                max(1, self.stats["total_calls"]), 0
            )
        }
    
    def _get_session_duration(self) -> float:
        """Calculate session duration in hours"""
        try:
            start = datetime.fromisoformat(self.stats["session_start"])
            duration = datetime.now() - start
            return round(duration.total_seconds() / 3600, 2)
        except:
            return 0.0
    
    def export_csv(self) -> str:
        """
        Export usage data to CSV format
        
        Returns:
            CSV string ready for download
        """
        import csv
        from io import StringIO
        
        output = StringIO()
        writer = csv.writer(output)
        
        # Header
        writer.writerow(["IM Creator - Anthropic API Usage Report"])
        writer.writerow(["Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
        writer.writerow([])
        
        # Summary section
        writer.writerow(["SUMMARY"])
        writer.writerow(["Total API Calls", self.stats["total_calls"]])
        writer.writerow(["Total Input Tokens", self.stats["total_input_tokens"]])
        writer.writerow(["Total Output Tokens", self.stats["total_output_tokens"]])
        writer.writerow(["Total Tokens", self.stats["total_input_tokens"] + self.stats["total_output_tokens"]])
        writer.writerow(["Total Cost (USD)", f"${self.stats['total_cost_usd']:.6f}"])
        writer.writerow(["Session Start", self.stats["session_start"]])
        writer.writerow(["Session Duration (hours)", self._get_session_duration()])
        writer.writerow([])
        
        # By-purpose breakdown
        writer.writerow(["BREAKDOWN BY PURPOSE"])
        writer.writerow(["Purpose", "Calls", "Tokens", "Cost (USD)"])
        for purpose, stats in self.stats["by_purpose"].items():
            writer.writerow([
                purpose,
                stats["calls"],
                stats["tokens"],
                f"${stats['cost']:.6f}"
            ])
        writer.writerow([])
        
        # By-model breakdown
        writer.writerow(["BREAKDOWN BY MODEL"])
        writer.writerow(["Model", "Calls", "Tokens", "Cost (USD)"])
        for model, stats in self.stats["by_model"].items():
            writer.writerow([
                model,
                stats["calls"],
                stats["tokens"],
                f"${stats['cost']:.6f}"
            ])
        writer.writerow([])
        
        # Recent calls (last 100)
        writer.writerow(["RECENT CALLS (Last 100)"])
        writer.writerow(["Timestamp", "Model", "Purpose", "Input Tokens", "Output Tokens", "Total Tokens", "Cost (USD)"])
        for call in self.stats["calls"][-100:]:
            writer.writerow([
                call["timestamp"],
                call.get("model_name", call["model"]),
                call["purpose"],
                call["input_tokens"],
                call["output_tokens"],
                call["total_tokens"],
                f"${call['cost_usd']:.6f}"
            ])
        
        return output.getvalue()
    
    def reset(self):
        """Reset all usage statistics"""
        self.stats = {
            "total_calls": 0,
            "total_input_tokens": 0,
            "total_output_tokens": 0,
            "total_cost_usd": 0.0,
            "session_start": datetime.now().isoformat(),
            "calls": [],
            "by_purpose": {},
            "by_model": {}
        }
        self._save_stats()
        print("Usage statistics reset successfully")
    
    def get_recent_calls(self, limit: int = 10) -> List[Dict]:
        """
        Get most recent API calls
        
        Args:
            limit: Number of recent calls to return (default: 10)
        
        Returns:
            List of recent call records
        """
        return self.stats["calls"][-limit:]


# ============================================================================
# GLOBAL INSTANCE
# ============================================================================

# Create a global tracker instance that can be imported
_tracker = UsageTracker()


def get_tracker() -> UsageTracker:
    """
    Get the global usage tracker instance
    
    Returns:
        Global UsageTracker instance
    """
    return _tracker