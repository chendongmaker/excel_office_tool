from __future__ import annotations

"""应用启动模块。

上游：`main.py` 作为项目根入口导入本模块。
下游：调用 UI 层的 `run()` 创建 QApplication 和主窗口。
"""

from excel_helper.ui.main_window import run


if __name__ == "__main__":
    # 支持 `python -m excel_helper.app` 方式启动。
    run()
