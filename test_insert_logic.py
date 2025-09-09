#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from process_attachments import find_attachment_by_number
from docx import Document

def test_insert_logic():
    """测试新的段落插入逻辑"""
    print('=== 测试新的段落插入逻辑 ===')

    # 模拟生成的项目文档内容
    project_docs = {
        '总体描述': '1. 测试总体描述第一段\n2. 测试总体描述第二段\n3. 测试总体描述第三段',
        '项目建设目标': '1. 测试目标第一段\n2. 测试目标第二段\n3. 测试目标第三段',
        '项目建设必要性': '1. 测试必要性第一段\n2. 测试必要性第二段\n3. 测试必要性第三段',
        '存在问题': '1. 测试问题第一段\n2. 测试问题第二段\n3. 测试问题第三段'
    }

    path1 = find_attachment_by_number(1)
    if not path1:
        print('未找到附件1文件')
        return

    try:
        doc = Document(path1)
        
        section_title_mappings = {
            '1.1': '总体描述',
            '1.2': '项目建设目标', 
            '1.3': '项目建设必要性',
            '2.3': '存在问题'
        }
        
        updated_sections = []
        
        print('查找章节标题中的标识...')
        
        # 直接查找章节标题中的标识
        for i, paragraph in enumerate(doc.paragraphs):
            text = paragraph.text.strip()
            
            # 查找带有标识的章节标题（支持两种格式）
            for section_num, section_name in section_title_mappings.items():
                # 格式1: "1.1 总体描述（添加标识）"
                # 格式2: "总体描述（添加标识）"  
                if ((text.startswith(section_num) and section_name in text and ('添加标识' in text or '标识' in text or '（' in text)) or
                    (text == f"{section_name}（添加标识）" or text.startswith(f"{section_name}（") and "标识" in text)):
                    
                    print(f'找到带标识的章节：第{i}行 - {text}')
                    
                    if section_name in project_docs:
                        content = project_docs[section_name].strip()
                        
                        if content:
                            # 在章节标题后添加内容
                            lines = [line.strip() for line in content.split('\n') if line.strip()]
                            
                            # 使用更简单可靠的方法：直接在目标段落后面依次插入
                            target_para = paragraph
                            parent = target_para._element.getparent()
                            target_element = target_para._element
                            
                            # 正序插入每一行内容
                            insert_position = list(parent).index(target_element) + 1
                            
                            for line_idx, line in enumerate(lines):
                                # 创建新的段落元素
                                new_p = doc._body._element.makeelement('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
                                # 创建文本运行
                                new_r = doc._body._element.makeelement('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
                                new_t = doc._body._element.makeelement('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                                new_t.text = line
                                new_r.append(new_t)
                                new_p.append(new_r)
                                
                                # 插入到正确位置
                                parent.insert(insert_position + line_idx, new_p)
                            
                            updated_sections.append(section_name)
                            print(f'✅ 已在章节 {section_num} 后添加 {len(lines)} 行内容')
                            
                            # 提前退出，避免重复处理
                            break
                    break
        
        if updated_sections:
            doc.save(path1)
            print(f'✅ Word文档自动更新成功，已替换标注：{", ".join(updated_sections)}')
            print('💡 内容已精确插入到您标注的位置')
        else:
            print('❌ 未找到任何标识')
            
    except Exception as e:
        print(f'测试失败: {e}')
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    test_insert_logic()
