#!/usr/bin/env python3
"""
Start the Krishna Engineering Attendance Tool.

Usage:
  python run.py                  # runs on http://localhost:8000
  python run.py --port 9000      # custom port
  python run.py --host 0.0.0.0   # expose on network
"""
import sys, os, argparse
sys.path.insert(0, os.path.dirname(__file__))

import uvicorn

parser = argparse.ArgumentParser(description="Attendance Tool Server")
parser.add_argument("--host", default="127.0.0.1")
parser.add_argument("--port", type=int, default=8000)
parser.add_argument("--reload", action="store_true", default=False)
args = parser.parse_args()

print(f"""
╔══════════════════════════════════════════════════════╗
║     Krishna Engineering · Attendance Tool v1.0       ║
╠══════════════════════════════════════════════════════╣
║  URL   →  http://{args.host}:{args.port:<5}                    ║
║  Stop  →  Ctrl+C                                     ║
╚══════════════════════════════════════════════════════╝
""")

uvicorn.run(
    "backend.main:app",
    host=args.host,
    port=args.port,
    reload=args.reload
)
