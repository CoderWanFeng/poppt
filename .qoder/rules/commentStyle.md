
---
trigger: manual
description: "Python代码注释遵循Google Python Style Guide规范"
---

# Python注释规范

本项目的Python代码注释遵循Google Python Style Guide。

## 主要规范要点

### 1. 文档字符串(Docstrings)
- 使用三重双引号`"""`
- 模块、类、函数都应该有docstring
- 格式：一行摘要 + 空行 + 详细描述

### 2. 函数/方法注释
def function_name(param1, param2):
    """一行简短描述。

    详细描述（如需要）。

    Args:
        param1: 参数1的描述
        param2: 参数2的描述

    Returns:
        返回值描述

    Raises:
        ExceptionType: 异常描述
    """

### 3. 类注释
class ClassName:
    """类的简短描述。

    详细描述（如需要）。

    Attributes:
        attribute1: 属性1描述
        attribute2: 属性2描述
    """

### 4. 行内注释
- 使用`#`开头
- 与代码至少间隔2个空格
- 注释应该是完整的句子

### 5. TODO注释
# TODO(username): 具体的待办事项描述

参考文档：https://google.github.io/styleguide/pyguide.html
trigger: manual
---
