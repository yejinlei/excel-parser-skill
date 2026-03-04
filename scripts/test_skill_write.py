#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试SKILL入口点的写Excel功能
"""

from excel_parser import main
import json

# 测试数据
test_data = {
    "action": "write",
    "output_path": "test_skill_output.xlsx",
    "data": {
        "sheets": [
            {
                "name": "SKILL测试",
                "rows": [
                    ["标题"],
                    ["数据1", "数据2"],
                    ["数据3", "数据4"]
                ]
            }
        ]
    }
}

# 调用SKILL入口点
result = main(test_data)

# 打印结果
print(json.dumps(result, ensure_ascii=False, indent=2))

# 验证文件是否生成
import os
if os.path.exists("test_skill_output.xlsx"):
    print("\n文件已成功生成！")
else:
    print("\n文件生成失败！")
