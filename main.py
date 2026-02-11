#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AIDA64报告智能解析工具 - 主程序入口
作者: AI助手
版本: 1.0.0
"""

import sys
import os
from gui import AIDA64ParserApp

def main():
    """主函数"""
    app = AIDA64ParserApp()
    app.run()

if __name__ == "__main__":
    main()
