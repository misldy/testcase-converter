# Test Case Converter

Convert test cases between Excel and XMind formats.

## Installation
```bash
pip install testcase-converter
```

## Usage Examples

### As a CLI Tool:
```bash
# Excel to XMind
testcase-converter test_cases.xlsx

# XMind to Excel
testcase-converter test_cases.xmind
```

### As a Python Library:
```python
from testcase_converter import TestCaseConverter, ConversionType

# Auto-detect conversion type
converter = TestCaseConverter("input.xlsx")
converter.convert()

# Explicitly specify conversion type
converter = TestCaseConverter("input.xmind", ConversionType.XMIND_TO_EXCEL)
converter.convert()
```