#!/usr/bin/env python3
"""Invoice Renamer — launcher.

Полная логика находится в пакете invoice_renamer/.
Запуск:  python invoice_renamer.py [--headless] [--setup] [--config path]
"""
from invoice_renamer.__main__ import main

if __name__ == "__main__":
    main()
