import unittest
import os
import tempfile
from testcase_converter import TestCaseConverter, ConversionType

class TestConverter(unittest.TestCase):
    def setUp(self):
        self.test_dir = tempfile.TemporaryDirectory()
        self.input_excel = os.path.join(self.test_dir.name, "test.xlsx")
        self.input_xmind = os.path.join(self.test_dir.name, "test.xmind")
        
        # 创建测试用的Excel文件
        # 这里添加创建Excel文件的代码
        
        # 创建测试用的XMind文件
        # 这里添加创建XMind文件的代码
    
    def tearDown(self):
        self.test_dir.cleanup()
    
    def test_excel_to_xmind(self):
        converter = TestCaseConverter(self.input_excel)
        converter.convert()
        # 添加断言检查输出文件
        
    def test_xmind_to_excel(self):
        converter = TestCaseConverter(self.input_xmind)
        converter.convert()
        # 添加断言检查输出文件

if __name__ == '__main__':
    unittest.main()