"""项目根启动入口。

用户运行 `python main.py` 时会进入这里，再转到 `excel_helper.app.run()`。
"""

from excel_helper.app import run


if __name__ == "__main__":
    # 桌面应用主入口：创建 Qt 应用、显示主窗口、进入事件循环。
    run()
