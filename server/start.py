from __future__ import annotations

import os
import sys
from pathlib import Path

import uvicorn


if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from server.app import app

    port = int(os.environ.get("PORT", "10000"))
    print(f"Starting ORA report service on 0.0.0.0:{port}", flush=True)
    uvicorn.run(app, host="0.0.0.0", port=port, log_level="info")
