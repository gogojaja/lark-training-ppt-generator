# cover-skill 示例


# 示例 1: 基础使用
from {skill_path.name.replace('-', '_')} import SomeClass

obj = SomeClass()
result = obj.do_something()
print(result)

# 示例 2: 高级使用
from {skill_path.name.replace('-', '_')} import AdvancedClass

obj = AdvancedClass()
obj.configure(param1='value1')
obj.configure(param2='value2')
result = obj.process()
print(result)

# 示例 3: 错误处理
try:
    obj = SomeClass()
    result = obj.do_something()
except Exception as e:
    print(f'错误: {e}')
