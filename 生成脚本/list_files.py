import os
import sys

doc_dir = os.path.join(os.getcwd(), '生成产物', '新场景操作手册_拆分')

print(f'Doc dir: {doc_dir}')
print(f'Exists: {os.path.exists(doc_dir)}')

# 列出文件
files = os.listdir(doc_dir)
print(f'Files: {files}')

# 查找包含"预"的文件
for file in files:
    if '预' in file:
        print(f'Found: {file}')
        print(f'File name bytes: {file.encode("utf-8")}')
