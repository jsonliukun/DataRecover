"""``python -m doc_renamer`` 的最小可执行入口。

当 ``-m`` 指向一个包时，Python 会执行包内的 ``__main__.py``。这里不复制任何
参数或业务逻辑，只把控制权交给 :func:`doc_renamer.cli.main`，从而让模块方式与
``pyproject.toml`` 注册的 ``doc-renamer`` 控制台脚本保持同一行为。

``if __name__ == "__main__"`` 防止测试或工具仅导入本模块时意外解析调用进程的
命令行。``SystemExit`` 把 ``main`` 返回的 0、1 或 130 传给操作系统，PowerShell、
批处理或调度系统便可通过退出码判断成功、普通失败和用户中断。
"""

from .cli import main


if __name__ == "__main__":
    raise SystemExit(main())
