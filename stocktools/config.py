"""Settings live in config.json at the repo root (copy config.example.json).

Relative paths are resolved from the repo root, so the sample data works
out of the box.
"""
import json
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def load():
    for name in ("config.json", "config.example.json"):
        path = os.path.join(ROOT, name)
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                cfg = json.load(f)
            for key, value in cfg.get("paths", {}).items():
                if not os.path.isabs(value):
                    cfg["paths"][key] = os.path.join(ROOT, value)
            return cfg
    raise FileNotFoundError("No config.json or config.example.json found")
