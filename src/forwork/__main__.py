"""支持 ``python -m forwork`` 直接运行（等价于 forwork CLI）。"""

from forwork.cli import main

if __name__ == "__main__":
    raise SystemExit(main())
