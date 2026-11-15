"""exe 启动器：定位打包内的 app.py，起 Streamlit 服务并自动打开浏览器。

业务代码都在 app.py / core / ui 里，这里只负责「双击 exe 就能用」。
打包命令：python -m PyInstaller excel-toolbox.spec --clean --noconfirm
"""

from __future__ import annotations

import os
import socket
import sys
import threading
import webbrowser
from pathlib import Path

PORT_START = 8501
PORT_TRIES = 20
OPEN_DELAY = 2.5


def bundle_dir() -> Path:
    """单文件模式资源解压在 sys._MEIPASS；源码直接运行时用本文件所在目录。"""
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))


def pick_port() -> int:
    """从 8501 往上找一个没被占用的端口，都在用就退回 8501。"""
    for port in range(PORT_START, PORT_START + PORT_TRIES):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            try:
                sock.bind(("0.0.0.0", port))
            except OSError:
                continue
            return port
    return PORT_START


def main() -> None:
    root = bundle_dir()
    app = root / "app.py"
    if not app.is_file():
        raise SystemExit(f"启动失败：找不到 {app}")
    os.chdir(root)

    port = pick_port()
    url = f"http://127.0.0.1:{port}"
    timer = threading.Timer(OPEN_DELAY, webbrowser.open, args=(url,))
    timer.daemon = True
    timer.start()

    sys.argv = [
        "streamlit",
        "run",
        str(app),
        "--server.address=0.0.0.0",
        f"--server.port={port}",
        "--server.headless=true",
        "--server.fileWatcherType=none",
        "--server.maxUploadSize=50",
        "--browser.gatherUsageStats=false",
    ]
    from streamlit.web import cli

    cli.main()


if __name__ == "__main__":
    main()
