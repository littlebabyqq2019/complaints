import zipfile

filename = "D:/Program Files/投诉批办单管理系统/待打印/投诉批办单#C-26-2583.docx"
with zipfile.ZipFile(filename, 'r') as docx:
    content = docx.open('word/document.xml').read().decode('utf-8')
    print('✓ 文档包含 wp:anchor:', 'wp:anchor' in content)
    print('✓ 文档包含 spacing line=1:', 'w:line="1"' in content)
    
    # 统计段落数
    para_count = content.count('<w:p ')
    print(f'段落数量: {para_count}')

