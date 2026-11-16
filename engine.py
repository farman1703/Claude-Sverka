"""
Модуль сверки для приложения.

compare_excel.py использует библиотеку pandas. В приложении вместо неё
подставляется pandas_lite (чистый Python): скомпилированные файлы pandas и numpy
Windows может заблокировать. Сам compare_excel.py при этом не меняется.

Использование:  from engine import ce
"""

import sys

import pandas_lite

sys.modules["pandas"] = pandas_lite     # compare_excel.py получит его как pandas

import compare_excel as ce  # noqa: E402,F401
