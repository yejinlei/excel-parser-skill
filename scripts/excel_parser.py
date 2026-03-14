#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel内容解析脚本
支持多种Excel格式的内容提取，使用python-calamine库
"""

import os
import sys
from typing import Dict, Any, List, Optional
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()


def install_dependency(package):
    """自动安装缺失的依赖"""
    import subprocess
    print(f"正在安装依赖: {package}")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        print(f"依赖 {package} 安装成功")
        return True
    except subprocess.CalledProcessError as e:
        print(f"依赖 {package} 安装失败: {e}")
        return False


class ExcelParser:
    """Excel文件解析器"""
    
    def __init__(self):
        """初始化Excel解析器"""
        self.calamine = None
        self._init_engine()
        # 加载环境变量配置
        self.max_rows = int(os.getenv('EXCEL_MAX_ROWS', 100))
        self.keep_empty_rows = os.getenv('EXCEL_KEEP_EMPTY_ROWS', 'false').lower() == 'true'
    
    def _init_engine(self):
        """初始化解析引擎"""
        try:
            from python_calamine import CalamineWorkbook
            self.calamine = CalamineWorkbook
        except ImportError:
            print("Calamine依赖未安装，正在尝试自动安装...")
            if install_dependency("python-calamine"):
                try:
                    from python_calamine import CalamineWorkbook
                    self.calamine = CalamineWorkbook
                except ImportError:
                    raise Exception("Calamine依赖安装失败，请手动安装: pip install python-calamine")
            else:
                raise Exception("Calamine依赖安装失败，请手动安装: pip install python-calamine")
    
    def parse_excel(self, excel_path: str) -> Dict[str, Any]:
        """
        解析Excel文件内容
        
        Args:
            excel_path: Excel文件路径
        
        Returns:
            包含解析结果的字典
        """
        result = {
            "sheets": [],
            "sheet_count": 0,
            "total_cells": 0,
            "engine": "python-calamine"
        }
        
        try:
            # 使用python-calamine解析Excel文件
            workbook = self.calamine.from_path(excel_path)
            
            # 获取所有工作表名称
            sheet_names = workbook.sheet_names
            result["sheet_count"] = len(sheet_names)
            
            # 解析每个工作表
            for idx, sheet_name in enumerate(sheet_names):
                sheet_data = {
                    "name": sheet_name,
                    "rows": [],
                    "row_count": 0,
                    "column_count": 0,
                    "merged_cells": []
                }
                
                # 获取工作表数据
                sheet = workbook.get_sheet_by_index(idx)
                if sheet:
                    # 读取所有行 - 使用to_python()方法获取数据
                    try:
                        # 尝试使用to_python()方法
                        data = sheet.to_python()
                        rows = []
                        for row in data:
                            # 处理行数据
                            row_data = []
                            for cell in row:
                                # 转换单元格值
                                if cell is None:
                                    row_data.append("")
                                elif isinstance(cell, str):
                                    row_data.append(cell.strip())
                                else:
                                    row_data.append(str(cell))
                            rows.append(row_data)
                    except AttributeError:
                        # 如果to_python()不可用，尝试直接访问
                        rows = []
                        # 获取工作表的行数和列数
                        try:
                            # 尝试遍历所有单元格
                            row_idx = 0
                            while True:
                                try:
                                    row_data = []
                                    col_idx = 0
                                    while True:
                                        try:
                                            cell = sheet.get_cell(row_idx, col_idx)
                                            if cell is None:
                                                row_data.append("")
                                            elif isinstance(cell, str):
                                                row_data.append(cell.strip())
                                            else:
                                                row_data.append(str(cell))
                                            col_idx += 1
                                        except:
                                            break
                                    if row_data:
                                        rows.append(row_data)
                                    row_idx += 1
                                except:
                                    break
                        except Exception as e2:
                            print(f"直接访问单元格失败: {e2}")
                            raise
                    
                    # 计算行数和列数
                    if rows:
                        sheet_data["rows"] = rows
                        sheet_data["row_count"] = len(rows)
                        sheet_data["column_count"] = len(rows[0]) if rows[0] else 0
                        result["total_cells"] += sheet_data["row_count"] * sheet_data["column_count"]
                
                result["sheets"].append(sheet_data)
            
        except Exception as e:
            # 如果python-calamine失败，尝试使用xlwings
            print(f"python-calamine解析失败，尝试使用xlwings: {e}")
            result = self._parse_with_fallback(excel_path)
        
        return result
    
    def _parse_with_fallback(self, excel_path: str) -> Dict[str, Any]:
        """
        使用xlwings作为备选解析方案
        """
        result = {
            "sheets": [],
            "sheet_count": 0,
            "total_cells": 0,
            "engine": "xlwings"
        }
        
        try:
            import xlwings as xw
            
            # 使用不可见模式打开Excel
            app = xw.App(visible=False)
            try:
                workbook = app.books.open(excel_path)
                
                # 获取所有工作表
                sheet_names = [sheet.name for sheet in workbook.sheets]
                result["sheet_count"] = len(sheet_names)
                
                # 解析每个工作表
                for sheet in workbook.sheets:
                    sheet_data = {
                        "name": sheet.name,
                        "rows": [],
                        "row_count": 0,
                        "column_count": 0
                    }
                    
                    # 获取工作表数据
                    rows = []
                    
                    # 获取已使用的范围
                    used_range = sheet.used_range
                    if used_range:
                        # 获取行数和列数
                        row_count = used_range.rows.count
                        col_count = used_range.columns.count
                        
                        # 读取所有数据
                        for row_idx in range(row_count):
                            row_data = []
                            for col_idx in range(col_count):
                                cell = sheet.cells(row_idx + 1, col_idx + 1)
                                value = cell.value
                                
                                # 转换单元格值
                                if value is None:
                                    row_data.append("")
                                elif isinstance(value, str):
                                    row_data.append(value.strip())
                                else:
                                    row_data.append(str(value))
                                rows.append(row_data)
                    
                    # 计算行数和列数
                    if rows:
                        sheet_data["rows"] = rows
                        sheet_data["row_count"] = len(rows)
                        sheet_data["column_count"] = len(rows[0]) if rows[0] else 0
                        result["total_cells"] += sheet_data["row_count"] * sheet_data["column_count"]
                    
                    result["sheets"].append(sheet_data)
                
                # 关闭工作簿
                workbook.close()
            finally:
                # 退出Excel应用
                app.quit()
            
        except ImportError:
            print("xlwings依赖未安装，正在尝试自动安装...")
            if install_dependency("xlwings"):
                # 重新导入并处理
                import xlwings as xw
                
                # 使用不可见模式打开Excel
                app = xw.App(visible=False)
                try:
                    workbook = app.books.open(excel_path)
                    
                    # 获取所有工作表
                    sheet_names = [sheet.name for sheet in workbook.sheets]
                    result["sheet_count"] = len(sheet_names)
                    
                    # 解析每个工作表
                    for sheet in workbook.sheets:
                        sheet_data = {
                            "name": sheet.name,
                            "rows": [],
                            "row_count": 0,
                            "column_count": 0
                        }
                        
                        # 获取工作表数据
                        rows = []
                        
                        # 获取已使用的范围
                        used_range = sheet.used_range
                        if used_range:
                            # 获取行数和列数
                            row_count = used_range.rows.count
                            col_count = used_range.columns.count
                            
                            # 读取所有数据
                            for row_idx in range(row_count):
                                row_data = []
                                for col_idx in range(col_count):
                                    cell = sheet.cells(row_idx + 1, col_idx + 1)
                                    value = cell.value
                                    
                                    # 转换单元格值
                                    if value is None:
                                        row_data.append("")
                                    elif isinstance(value, str):
                                        row_data.append(value.strip())
                                    else:
                                        row_data.append(str(value))
                                    rows.append(row_data)
                        
                        # 计算行数和列数
                        if rows:
                            sheet_data["rows"] = rows
                            sheet_data["row_count"] = len(rows)
                            sheet_data["column_count"] = len(rows[0]) if rows[0] else 0
                            result["total_cells"] += sheet_data["row_count"] * sheet_data["column_count"]
                        
                        result["sheets"].append(sheet_data)
                    
                    # 关闭工作簿
                    workbook.close()
                finally:
                    # 退出Excel应用
                    app.quit()
            else:
                raise Exception("xlwings依赖安装失败，请手动安装: pip install xlwings")
        
        except Exception as e:
            raise Exception(f"Excel解析失败: {str(e)}")
        
        return result
    
    def parse_excel_to_text(self, excel_path: str) -> str:
        """
        将Excel文件解析为文本格式
        
        Args:
            excel_path: Excel文件路径
        
        Returns:
            解析后的文本
        """
        try:
            result = self.parse_excel(excel_path)
            
            text_parts = []
            text_parts.append(f"=== Excel文件: {os.path.basename(excel_path)} ===")
            text_parts.append(f"工作表数量: {result['sheet_count']}")
            text_parts.append(f"总单元格数: {result['total_cells']}")
            text_parts.append(f"使用引擎: {result['engine']}")
            text_parts.append("")
            
            # 处理每个工作表
            for sheet in result['sheets']:
                text_parts.append(f"--- 工作表: {sheet['name']} ---")
                text_parts.append(f"行数: {sheet['row_count']}, 列数: {sheet['column_count']}")
                if 'merged_cells' in sheet and sheet['merged_cells']:
                    text_parts.append(f"合并单元格数量: {len(sheet['merged_cells'])}")
                text_parts.append("")
                
                # 输出数据，使用配置的最大行数限制
                max_rows = min(sheet['row_count'], self.max_rows)
                for i, row in enumerate(sheet['rows'][:max_rows]):
                    # 根据配置决定是否保留空行
                    if self.keep_empty_rows or any(cell.strip() for cell in row):
                        row_text = "\t".join(row)
                        text_parts.append(f"{i+1}: {row_text}")
                
                # 如果有更多行，提示
                if sheet['row_count'] > self.max_rows:
                    text_parts.append(f"... 还有 {sheet['row_count'] - self.max_rows} 行未显示")
                
                text_parts.append("")
            
            return "\n".join(text_parts)
            
        except Exception as e:
            return f"【Excel文件处理失败: {str(e)}】"
    
    def write_excel(self, output_path: str, data: Dict[str, Any]) -> bool:
        """
        写入Excel文件
        
        Args:
            output_path: 输出文件路径
            data: 要写入的数据，格式如下：
                {
                    "sheets": [
                        {
                            "name": "Sheet1",
                            "rows": [["A1", "B1"], ["A2", "B2"]],
                            "merged_cells": [
                                {
                                    "start_row": 0,
                                    "end_row": 0,
                                    "start_col": 0,
                                    "end_col": 1
                                }
                            ]
                        }
                    ]
                }
        
        Returns:
            是否写入成功
        """
        try:
            import xlwings as xw
            
            # 使用不可见模式打开Excel
            app = xw.App(visible=False)
            try:
                # 创建新工作簿
                workbook = app.books.add()
                
                # 处理数据
                sheets_data = data.get('sheets', [])
                
                if not sheets_data:
                    # 如果没有提供数据，使用默认工作表
                    workbook.sheets[0].name = "Sheet1"
                else:
                    # 删除默认工作表
                    if len(workbook.sheets) > 1:
                        workbook.sheets[0].delete()
                    
                    # 创建新工作表
                    for sheet_data in sheets_data:
                        sheet_name = sheet_data.get('name', 'Sheet')
                        rows = sheet_data.get('rows', [])
                        merged_cells = sheet_data.get('merged_cells', [])
                        
                        # 检查工作表是否存在
                        sheet = None
                        for existing_sheet in workbook.sheets:
                            if existing_sheet.name == sheet_name:
                                sheet = existing_sheet
                                break
                        
                        if not sheet:
                            # 创建新工作表
                            sheet = workbook.sheets.add(after=workbook.sheets[-1])
                            sheet.name = sheet_name
                        
                        # 写入数据
                        for row_idx, row in enumerate(rows):
                            for col_idx, cell_value in enumerate(row):
                                sheet.range((row_idx + 1, col_idx + 1)).value = cell_value
                        
                        # 处理合并单元格
                        for merged_cell in merged_cells:
                            start_row = merged_cell.get('start_row', 0) + 1
                            end_row = merged_cell.get('end_row', 0) + 1
                            start_col = merged_cell.get('start_col', 0) + 1
                            end_col = merged_cell.get('end_col', 0) + 1
                            
                            if start_row <= end_row and start_col <= end_col:
                                start_cell = (start_row, start_col)
                                end_cell = (end_row, end_col)
                                sheet.range(start_cell, end_cell).api.merge()
                
                # 保存文件
                workbook.save(output_path)
                workbook.close()
                
                return True
                
            finally:
                # 退出Excel应用
                app.quit()
            
        except Exception as e:
            print(f"写入Excel文件失败: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def update_excel(self, excel_path: str, data: Dict[str, Any]) -> bool:
        """
        更新现有Excel文件
        
        Args:
            excel_path: 现有Excel文件路径
            data: 要更新的数据，格式同write_excel
        
        Returns:
            是否更新成功
        """
        try:
            import xlwings as xw
            
            # 使用不可见模式打开Excel
            app = xw.App(visible=False)
            try:
                # 打开现有工作簿
                workbook = app.books.open(excel_path)
                
                # 处理数据
                sheets_data = data.get('sheets', [])
                
                for sheet_data in sheets_data:
                    sheet_name = sheet_data.get('name')
                    rows = sheet_data.get('rows', [])
                    merged_cells = sheet_data.get('merged_cells', [])
                    
                    if sheet_name:
                        # 检查工作表是否存在
                        sheet = None
                        for existing_sheet in workbook.sheets:
                            if existing_sheet.name == sheet_name:
                                sheet = existing_sheet
                                break
                        
                        if not sheet:
                            # 创建新工作表
                            sheet = workbook.sheets.add(after=workbook.sheets[-1])
                            sheet.name = sheet_name
                        
                        # 清空现有数据（可选）
                        # sheet.clear()
                        
                        # 写入数据
                        for row_idx, row in enumerate(rows):
                            for col_idx, cell_value in enumerate(row):
                                sheet.range((row_idx + 1, col_idx + 1)).value = cell_value
                        
                        # 处理合并单元格
                        for merged_cell in merged_cells:
                            start_row = merged_cell.get('start_row', 0) + 1
                            end_row = merged_cell.get('end_row', 0) + 1
                            start_col = merged_cell.get('start_col', 0) + 1
                            end_col = merged_cell.get('end_col', 0) + 1
                            
                            if start_row <= end_row and start_col <= end_col:
                                start_cell = (start_row, start_col)
                                end_cell = (end_row, end_col)
                                sheet.range(start_cell, end_cell).api.merge()
                
                # 保存文件
                workbook.save()
                workbook.close()
                
                return True
                
            finally:
                # 退出Excel应用
                app.quit()
            
        except Exception as e:
            print(f"更新Excel文件失败: {e}")
            import traceback
            traceback.print_exc()
            return False


def process_excel(excel_path: str) -> Dict[str, Any]:
    """
    处理Excel文件的主函数
    
    Args:
        excel_path: Excel文件路径
    
    Returns:
        包含text、sheets和sheet_count的字典
    """
    parser = ExcelParser()
    excel_data = parser.parse_excel(excel_path)
    text = parser.parse_excel_to_text(excel_path)
    
    return {
        "text": text,
        "sheets": excel_data["sheets"],
        "sheet_count": excel_data["sheet_count"],
        "total_cells": excel_data["total_cells"],
        "engine": excel_data["engine"]
    }


def write_excel_file(output_path: str, data: Dict[str, Any]) -> Dict[str, Any]:
    """
    写入Excel文件的主函数
    
    Args:
        output_path: 输出文件路径
        data: 要写入的数据
    
    Returns:
        包含success和message的字典
    """
    parser = ExcelParser()
    success = parser.write_excel(output_path, data)
    
    if success:
        return {
            "success": True,
            "message": f"Excel文件已成功写入: {output_path}",
            "file_path": output_path
        }
    else:
        return {
            "success": False,
            "error": f"写入Excel文件失败: {output_path}"
        }


def update_excel_file(excel_path: str, data: Dict[str, Any]) -> Dict[str, Any]:
    """
    更新Excel文件的主函数
    
    Args:
        excel_path: 现有Excel文件路径
        data: 要更新的数据
    
    Returns:
        包含success和message的字典
    """
    parser = ExcelParser()
    success = parser.update_excel(excel_path, data)
    
    if success:
        return {
            "success": True,
            "message": f"Excel文件已成功更新: {excel_path}",
            "file_path": excel_path
        }
    else:
        return {
            "success": False,
            "error": f"更新Excel文件失败: {excel_path}"
        }


def main(input_data: Dict[str, Any] = None) -> Dict[str, Any]:
    """SKILL 入口点"""
    if input_data is None:
        input_data = {}
    
    # 确定操作类型
    action = input_data.get('action', 'read').lower()
    
    if action == 'read':
        # 读取Excel文件
        # 支持多种参数名称，提高兼容性
        excel_path = input_data.get('file_path', '') or input_data.get('file', '')
        if not excel_path:
            return {"success": False, "error": "Excel file path is required"}
        
        try:
            result = process_excel(excel_path)
            return {"success": True, "result": result}
        except Exception as e:
            return {"success": False, "error": f"Excel processing failed: {str(e)}"}
    
    elif action == 'write':
        # 写入Excel文件
        output_path = input_data.get('output_path', '') or input_data.get('file_path', '') or input_data.get('file', '')
        data = input_data.get('data', {})
        
        if not output_path:
            return {"success": False, "error": "Output file path is required"}
        
        if not data.get('sheets'):
            return {"success": False, "error": "Excel data is required"}
        
        try:
            result = write_excel_file(output_path, data)
            if result['success']:
                return {"success": True, "result": result}
            else:
                return {"success": False, "error": result.get('error', 'Write failed')}
        except Exception as e:
            return {"success": False, "error": f"Excel write failed: {str(e)}"}
    
    elif action == 'update':
        # 更新Excel文件
        excel_path = input_data.get('file_path', '') or input_data.get('file', '')
        data = input_data.get('data', {})
        
        if not excel_path:
            return {"success": False, "error": "Excel file path is required"}
        
        if not data.get('sheets'):
            return {"success": False, "error": "Excel data is required"}
        
        try:
            result = update_excel_file(excel_path, data)
            if result['success']:
                return {"success": True, "result": result}
            else:
                return {"success": False, "error": result.get('error', 'Update failed')}
        except Exception as e:
            return {"success": False, "error": f"Excel update failed: {str(e)}"}
    
    else:
        return {"success": False, "error": f"Invalid action: {action}. Supported actions: read, write, update"}


if __name__ == "__main__":
    # 测试代码
    if len(sys.argv) > 1:
        action = sys.argv[1]
        
        if action == 'read':
            # 测试读取Excel文件
            if len(sys.argv) > 2:
                excel_path = sys.argv[2]
            else:
                print("使用方法: python excel_parser.py read <excel_file_path>")
                sys.exit(1)
            
            if not os.path.exists(excel_path):
                print(f"文件不存在: {excel_path}")
                sys.exit(1)
            
            try:
                result = process_excel(excel_path)
                print(f"Excel解析完成，共 {result['sheet_count']} 个工作表")
                print(f"总单元格数: {result['total_cells']}")
                print(f"使用引擎: {result['engine']}")
                print("\n解析结果:")
                print(result['text'])
            except Exception as e:
                print(f"处理失败: {e}")
                sys.exit(1)
        
        elif action == 'write':
            # 测试写入Excel文件
            if len(sys.argv) > 2:
                output_path = sys.argv[2]
            else:
                print("使用方法: python excel_parser.py write <output_file_path>")
                sys.exit(1)
            
            try:
                test_data = {
                    "sheets": [
                        {
                            "name": "测试工作表",
                            "rows": [
                                ["姓名", "年龄", "城市"],
                                ["张三", 25, "北京"],
                                ["李四", 30, "上海"],
                                ["王五", 35, "广州"]
                            ],
                            "merged_cells": [
                                {
                                    "start_row": 0,
                                    "end_row": 0,
                                    "start_col": 0,
                                    "end_col": 2
                                }
                            ]
                        }
                    ]
                }
                
                result = write_excel_file(output_path, test_data)
                if result['success']:
                    print(f"Excel文件写入成功: {result['file_path']}")
                else:
                    print(f"Excel文件写入失败: {result['error']}")
            except Exception as e:
                print(f"写入失败: {e}")
                sys.exit(1)
        
        elif action == 'update':
            # 测试更新Excel文件
            if len(sys.argv) > 3:
                excel_path = sys.argv[2]
                output_path = sys.argv[3]
            else:
                print("使用方法: python excel_parser.py update <input_file_path> <output_file_path>")
                sys.exit(1)
            
            if not os.path.exists(excel_path):
                print(f"文件不存在: {excel_path}")
                sys.exit(1)
            
            try:
                # 先读取文件
                parser = ExcelParser()
                excel_data = parser.parse_excel(excel_path)
                
                # 修改数据
                if excel_data['sheets']:
                    # 在第一个工作表中添加一行
                    excel_data['sheets'][0]['rows'].append(["赵六", 40, "深圳"])
                
                # 写入到新文件
                result = write_excel_file(output_path, excel_data)
                if result['success']:
                    print(f"Excel文件更新成功: {result['file_path']}")
                else:
                    print(f"Excel文件更新失败: {result['error']}")
            except Exception as e:
                print(f"更新失败: {e}")
                sys.exit(1)
        
        else:
            print("支持的操作: read, write, update")
            sys.exit(1)
    else:
        print("使用方法:")
        print("  读取Excel: python excel_parser.py read <excel_file_path>")
        print("  写入Excel: python excel_parser.py write <output_file_path>")
        print("  更新Excel: python excel_parser.py update <input_file_path> <output_file_path>")
        sys.exit(1)
