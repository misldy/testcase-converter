import unittest
import os
import tempfile
from openpyxl import Workbook
from testcase_converter import TestCaseConverter, ConversionType

class TestConverter(unittest.TestCase):
    def setUp(self):
        self.test_dir = tempfile.TemporaryDirectory()
        
        # 创建符合要求的Excel测试文件（7列数据）
        self.input_excel = os.path.join(self.test_dir.name, "test.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.append(["模块路径", "用例名称", "前置条件", "执行步骤", "预期结果", "车型", "优先级"])
        ws.append(["模块1→模块2", "TC001", "无", "步骤1\n步骤2", "预期结果1\n预期结果2", "ModelX", "高"])
        wb.save(self.input_excel)
        
        # 创建更完整的XMind测试文件
        self.input_xmind = os.path.join(self.test_dir.name, "test.xmind")
        xmind_content = """<?xml version="1.0" encoding="UTF-8"?>
        <xmap-content><sheet><topic><title>Root</title>
        <topic><title>模块1→模块2→TC001</title>
        <notes><![CDATA[【前置条件】无
        【执行步骤】步骤1\n步骤2
        【预期结果】预期结果1\n预期结果2
        【车型】ModelX
        【优先级】高]]></notes></topic></topic></sheet></xmap-content>"""
        with open(self.input_xmind, 'w', encoding='utf-8') as f:
            f.write(xmind_content)
    
    def tearDown(self):
        # 确保所有资源都已关闭
        if hasattr(self, 'converter'):
            self.converter.close()
        self.test_dir.cleanup()
    
    def test_excel_to_xmind(self):
        self.converter = TestCaseConverter(self.input_excel, ConversionType.EXCEL_TO_XMIND)
        self.converter.convert()
        output_path = os.path.join(self.test_dir.name, f"{os.path.splitext(os.path.basename(self.input_excel))[0]}.xmind")
        self.assertTrue(os.path.exists(output_path))
    
    def test_xmind_to_excel(self):
        self.converter = TestCaseConverter(self.input_xmind, ConversionType.XMIND_TO_EXCEL)
        self.converter.convert()
        output_path = os.path.join(self.test_dir.name, f"{os.path.splitext(os.path.basename(self.input_xmind))[0]}.xlsx")
        self.assertTrue(os.path.exists(output_path))
        # 确保Excel文件至少有一个可见的工作表
        from openpyxl import load_workbook
        wb = load_workbook(output_path)
        self.assertGreaterEqual(len(wb.sheetnames), 1)

if __name__ == '__main__':
    unittest.main()