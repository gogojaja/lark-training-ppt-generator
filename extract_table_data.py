import os
import zipfile
import xml.etree.ElementTree as ET
import json

doc_dir = os.path.join(os.getcwd(), '生成产物', '新场景操作手册_拆分')

# 读取"预约查询.docx"的表格数据
for file in os.listdir(doc_dir):
    if '预约查询' in file and file.endswith('.docx'):
        print(f'Reading: {file}')

        doc_path = os.path.join(doc_dir, file.encode('utf-8').decode('utf-8'))
        with zipfile.ZipFile(doc_path, 'r') as z:
            xml_content = z.read('word/document.xml')
            root = ET.fromstring(xml_content)

            # 定义命名空间
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

            # 查找所有表格
            tables = root.findall('.//w:tbl', ns)

            print(f'Tables: {len(tables)}')

            for table_idx, table in enumerate(tables):
                # 提取表格数据
                rows = table.findall('.//w:tr', ns)
                table_data = []

                for row in rows:
                    cells = []
                    for cell in row.findall('.//w:tc', ns):
                        paragraphs = cell.findall('.//w:p', ns)
                        texts = []
                        for p in paragraphs:
                            runs = p.findall('.//w:r', ns)
                            for run in runs:
                                text_runs = run.findall('.//w:t', ns)
                                for text_run in text_runs:
                                    texts.append(text_run.text)
                        cells.append(''.join(texts))
                    table_data.append(cells)

                # 保存为 JSON
                json_data = {
                    'title': '预约查询 - 高拍仪品牌配置表',
                    'headers': table_data[0],
                    'rows': table_data[1:]
                }

                output_path = os.path.join(os.getcwd(), '生成产物', '表格', '预约查询.json')
                os.makedirs(os.path.dirname(output_path), exist_ok=True)

                with open(output_path, 'w', encoding='utf-8') as f:
                    json.dump(json_data, f, ensure_ascii=False, indent=2)

                print(f'Saved to: {output_path}')
                # print(f'Table data: {json.dumps(json_data, ensure_ascii=False)}')
