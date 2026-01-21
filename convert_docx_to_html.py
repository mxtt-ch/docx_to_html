#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DOCX转HTML核心转换类
基于原生XML解析，支持DOM树遍历、样式继承、列表编号、图片标注合成等功能
完全基于XML动态生成CSS样式，无硬编码样式
"""

import os
import re
import sys
import math
from io import BytesIO
from docx import Document
from docx.oxml.ns import qn
from PIL import Image as PILImage, ImageDraw, ImageFont
from html import escape
from tqdm import tqdm
import xml.etree.ElementTree as ET

from util import DocxUtils


class DocxToHTMLConverter:
    """DOCX转HTML转换器 - 核心类"""
    
    def __init__(self, docx_path, output_html_path, images_dir, xml_dir, log_file, title=None):
        """
        初始化转换器

        Args:
            docx_path: 输入的DOCX文件路径
            output_html_path: 输出的HTML文件路径
            images_dir: 图片保存目录
            xml_dir: XML文件保存目录
            log_file: 日志文件路径
        """
        self.docx_path = docx_path
        self.output_html_path = output_html_path
        self.images_dir = images_dir
        self.xml_dir = xml_dir

        # 加载文档
        self.doc = Document(docx_path)
        self.image_counter = 0

        # 打开日志文件
        self.log_file = open(log_file, 'w', encoding='utf-8')

        self.doc_title = title if title else "ScholarHub数智服务平台使用指南（2025）"

        # 目录相关
        self.headings_map = {}  # 标题映射: {heading_text: heading_id}
        self.paragraph_ids = {} # 段落ID映射: {paragraph_element: heading_id}
        self.toc_titles = set()  # 目录条目集合，用于辅助识别标题

        # 列表编号追踪
        self.list_counters = {}  # 格式: {numId: {ilvl: counter}}

        # 资源映射
        self.all_parts = {}  # 所有parts: {part_name: part}
        self.all_xml_parts = {}  # XML文件映射: {part_name: part}

        # 样式定义和继承链
        self.style_definitions = {}
        self.style_inheritance = {}
        self.numbering_definitions = {}
        self.table_styles = {}

        # 记录开始
        self.log_file.write(f"DOCX转HTML转换器 - 初始化\n")
        self.log_file.write(f"{'='*80}\n")
        self.log_file.write(f"输入文件: {docx_path}\n")
        self.log_file.write(f"输出HTML: {output_html_path}\n")
        self.log_file.write(f"图片目录: {images_dir}\n")
        self.log_file.write(f"XML目录: {xml_dir}\n")
        self.log_file.write(f"文档标题: {self.doc_title}\n")
        self.log_file.write(f"{'='*80}\n")

        # 加载所有资源
        self._load_all_resources()

        # 加载样式和编号定义
        self._load_style_definitions()
        self._load_numbering_definitions()
        self._load_table_styles()

        # 初始化缩放因子，使内容区宽度映射到1200px
        try:
            self._init_scale_to_1200()
        except Exception as e:
            self.log_file.write(f"初始化缩放因子失败，使用1.0: {e}\n")
            DocxUtils.SCALE = 1.0

    def _init_scale_to_1200(self):
        """计算DOCX内容区宽度并设置全局缩放因子到1200px"""
        def emu_to_px_no_scale(emu):
            return int(emu * 96 / 914400) if emu else 0
        try:
            section = self.doc.sections[0]
            page_w = getattr(section, 'page_width', None)
            left_m = getattr(section, 'left_margin', None)
            right_m = getattr(section, 'right_margin', None)
            content_emu = int(page_w) - int(left_m) - int(right_m) if page_w and left_m is not None and right_m is not None else None
            if content_emu and content_emu > 0:
                content_px = emu_to_px_no_scale(int(content_emu))
                scale = 1200 / content_px if content_px > 0 else 1.0
                # 限制缩放范围避免异常
                if scale < 0.5:
                    scale = 0.5
                if scale > 2.5:
                    scale = 2.5
                DocxUtils.SCALE = scale
                self.log_file.write(f"内容区宽度: {content_px}px -> 目标1200px, SCALE={scale:.4f}\n")
            else:
                DocxUtils.SCALE = 1.0
                self.log_file.write("无法获取内容区宽度，SCALE=1.0\n")
        except Exception as e:
            DocxUtils.SCALE = 1.0
            self.log_file.write(f"缩放计算异常，SCALE=1.0: {e}\n")

    def _load_all_resources(self):
        """加载docx包中的所有资源文件（仅XML文件，不包括图片）"""
        self.log_file.write(f"\n{'='*80}\n")
        self.log_file.write(f"加载资源文件...\n")
        self.log_file.write(f"{'='*80}\n")
        
        # 创建图片目录
        os.makedirs(self.images_dir, exist_ok=True)
        
        # 加载所有parts（包括图片），用于后续按需处理
        self.all_parts = {}  # {part_name: part}
        self.log_file.write(f"  遍历所有parts...\n")
        
        for part in self.doc.part.package.parts:
            part_name = part.partname
            self.all_parts[part_name] = part
            
            # 只保存XML文件到xml目录
            if part_name.endswith('.xml'):
                self.log_file.write(f"  加载XML: {part_name}\n")
                self.all_xml_parts[part_name] = part
                # 保存XML文件到xml目录
                xml_filename = part_name.lstrip('/').replace('/', '_')
                xml_path = os.path.join(self.xml_dir, xml_filename)
                os.makedirs(os.path.dirname(xml_path), exist_ok=True)
                with open(xml_path, 'wb') as f:
                    f.write(part.blob)
            elif not part_name.endswith('.xml'):
                self.log_file.write(f"  找到资源: {part_name}\n")
            
        self.log_file.write(f"\n加载完成: {len(self.all_xml_parts)} 个XML文件, {len(self.all_parts)} 个总parts\n")
    
    def _load_numbering_definitions(self):
        """加载numbering.xml中的列表定义"""
        self.log_file.write(f"\n{'='*80}\n")
        self.log_file.write(f"加载列表定义...\n")
        self.log_file.write(f"{'='*80}\n")
        
        numbering_defs = {}
        
        try:
            # 获取numbering.xml部分
            numbering_part = self.doc.part.numbering_part
            if numbering_part is None:
                self.log_file.write("未找到numbering.xml文件\n")
                return numbering_defs
            
            # 解析numbering.xml
            numbering_xml = numbering_part.element
            self.log_file.write(f"numbering.xml加载成功\n")
            
            # 查找所有num定义
            for num in numbering_xml.findall(f'.//{DocxUtils.W_NS}num'):
                numId = num.get(f'{DocxUtils.W_NS}numId')
                if numId:
                    numId = int(numId)
                    # 获取对应的abstractNum
                    abstractNumId_elem = num.find(f'{DocxUtils.W_NS}abstractNumId')
                    if abstractNumId_elem is not None:
                        abstractNumId = int(abstractNumId_elem.get(f'{DocxUtils.W_NS}val'))
                        
                        # 在abstractNum中查找对应的定义
                        for abstractNum in numbering_xml.findall(f'.//{DocxUtils.W_NS}abstractNum'):
                            if int(abstractNum.get(f'{DocxUtils.W_NS}abstractNumId')) == abstractNumId:
                                # 解析每个级别的列表定义
                                levels = {}
                                for lvl in abstractNum.findall(f'{DocxUtils.W_NS}lvl'):
                                    ilvl = int(lvl.get(f'{DocxUtils.W_NS}ilvl'))
                                    
                                    # 判断是编号还是项目符号
                                    is_bullet = False
                                    bullet_char = '•'
                                    numFmt = None
                                    
                                    # 检查编号格式
                                    numFmt_elem = lvl.find(f'{DocxUtils.W_NS}numFmt')
                                    if numFmt_elem is not None:
                                        numFmt = numFmt_elem.get(f'{DocxUtils.W_NS}val')
                                        if numFmt == 'bullet':
                                            is_bullet = True
                                    
                                    # 检查项目符号元素
                                    buChar = lvl.find(f'{DocxUtils.W_NS}buChar')
                                    buAutoNum = lvl.find(f'{DocxUtils.W_NS}buAutoNum')
                                    buBlip = lvl.find(f'{DocxUtils.W_NS}buBlip')
                                    
                                    if buChar is not None or buAutoNum is not None or buBlip is not None:
                                        is_bullet = True
                                        if buChar is not None:
                                            bullet_char = buChar.get(f'{DocxUtils.W_NS}char', '•')
                                    
                                    # 获取列表标记颜色
                                    bullet_color = None
                                    lvl_rPr = lvl.find(f'{DocxUtils.W_NS}rPr')
                                    if lvl_rPr is not None:
                                        color = lvl_rPr.find(f'{DocxUtils.W_NS}color')
                                        if color is not None:
                                            color_val = color.get(f'{DocxUtils.W_NS}val')
                                            if color_val and len(color_val) == 6:
                                                bullet_color = f"#{color_val}"
                                    
                                    # 获取列表标记字体大小
                                    bullet_size = None
                                    if lvl_rPr is not None:
                                        sz = lvl_rPr.find(f'{DocxUtils.W_NS}sz')
                                        if sz is not None:
                                            sz_val = sz.get(f'{DocxUtils.W_NS}val')
                                            if sz_val:
                                                bullet_size = int(sz_val) / 2
                                    
                                    levels[ilvl] = {
                                        'is_bullet': is_bullet,
                                        'bullet_char': bullet_char,
                                        'numFmt': numFmt,
                                        'bullet_color': bullet_color,
                                        'bullet_size': bullet_size
                                    }
                                
                                numbering_defs[numId] = levels
                                self.log_file.write(f"  加载列表定义 numId={numId}, abstractNumId={abstractNumId}, 级别数={len(levels)}\n")
                                break
            
        except Exception as e:
            self.log_file.write(f"加载numbering.xml失败: {e}\n")
        
        self.numbering_definitions = numbering_defs
        self.log_file.write(f"共加载 {len(numbering_defs)} 个列表定义\n")
    
    def _load_style_definitions(self):
        """加载styles.xml中的样式定义并构建继承链"""
        self.log_file.write(f"\n{'='*80}\n")
        self.log_file.write(f"加载样式定义...\n")
        self.log_file.write(f"{'='*80}\n")
        
        style_defs = {}
        
        try:
            styles_xml = self.doc.styles.element
            if styles_xml is None:
                self.log_file.write("未找到styles.xml文件\n")
                return style_defs
            
            self.log_file.write(f"styles.xml加载成功\n")
            
            # 查找所有样式定义
            for style in styles_xml.findall(f'.//{DocxUtils.W_NS}style'):
                style_id = style.get(f'{DocxUtils.W_NS}styleId')
                style_type = style.get(f'{DocxUtils.W_NS}type', 'paragraph')
                
                if style_id:
                    style_def = {
                        'type': style_type,
                        'name': None,
                        'based_on': None,
                        'pPr': None,
                        'rPr': None
                    }
                    
                    # 获取样式名称
                    name_elem = style.find(f'{DocxUtils.W_NS}name')
                    if name_elem is not None:
                        style_def['name'] = name_elem.get(f'{DocxUtils.W_NS}val')
                    
                    # 获取基于的样式
                    based_on_elem = style.find(f'{DocxUtils.W_NS}basedOn')
                    if based_on_elem is not None:
                        style_def['based_on'] = based_on_elem.get(f'{DocxUtils.W_NS}val')
                    
                    # 获取段落属性
                    pPr = style.find(f'{DocxUtils.W_NS}pPr')
                    if pPr is not None:
                        style_def['pPr'] = pPr
                    
                    # 获取run属性
                    rPr = style.find(f'{DocxUtils.W_NS}rPr')
                    if rPr is not None:
                        style_def['rPr'] = rPr
                    
                    style_defs[style_id] = style_def
            
        except Exception as e:
            self.log_file.write(f"加载styles.xml失败: {e}\n")
        
        # 构建样式继承链
        self.style_inheritance = {}
        for style_id, style_def in style_defs.items():
            chain = [style_id]
            based_on = style_def.get('based_on')
            while based_on:
                chain.append(based_on)
                if based_on in style_defs:
                    based_on = style_defs[based_on].get('based_on')
                else:
                    break
            self.style_inheritance[style_id] = chain
        
        self.style_definitions = style_defs
        self.log_file.write(f"共加载 {len(style_defs)} 个样式定义\n")
        self.log_file.write(f"构建了 {len(self.style_inheritance)} 个样式继承链\n")
    
    def _load_table_styles(self):
        """加载表格样式定义"""
        self.log_file.write(f"\n{'='*80}\n")
        self.log_file.write(f"加载表格样式定义...\n")
        self.log_file.write(f"{'='*80}\n")
        
        table_styles = {}
        
        try:
            styles_xml = self.doc.styles.element
            if styles_xml is None:
                return table_styles
            
            for style in styles_xml.findall(f'.//{DocxUtils.W_NS}style'):
                style_id = style.get(f'{DocxUtils.W_NS}styleId')
                style_type = style.get(f'{DocxUtils.W_NS}type', 'paragraph')
                
                if style_id and style_type == 'table':
                    table_styles[style_id] = {
                        'tblPr': style.find(f'{DocxUtils.W_NS}tblPr'),
                        'tblStylePr': style.findall(f'.//{DocxUtils.W_NS}tblStylePr')
                    }
            
        except Exception as e:
            self.log_file.write(f"加载表格样式失败: {e}\n")
        
        self.table_styles = table_styles
        self.log_file.write(f"共加载 {len(table_styles)} 个表格样式\n")
    
    def _get_resolved_style_css(self, style_name, property_type='paragraph'):
        """
        获取样式及其所有父样式的合并CSS
        
        Args:
            style_name: 样式名称
            property_type: 样式类型 ('paragraph' 或 'run')
            
        Returns:
            合并后的CSS字符串
        """
        if style_name not in self.style_inheritance:
            return ""
        
        css_parts = []
        inheritance_chain = self.style_inheritance[style_name]
        
        # 按继承顺序从父到子遍历（先应用基础样式，后应用子样式）
        for ancestor_style_id in reversed(inheritance_chain):
            if ancestor_style_id in self.style_definitions:
                style_def = self.style_definitions[ancestor_style_id]
                if property_type == 'paragraph' and style_def['pPr']:
                    css = DocxUtils.get_paragraph_style_css(style_def['pPr'], self.log_file, apply_scale=False)
                    if css:
                        css_parts.append(css)
                elif property_type == 'run' and style_def['rPr']:
                    css = DocxUtils.get_run_style_css(style_def['rPr'], apply_scale=False)
                    if css:
                        css_parts.append(css)
        
        return DocxUtils.merge_css_with_priority(css_parts)
    
    def _get_list_number(self, numId, ilvl, numFmt=None):
        """
        获取列表项的编号
        
        Args:
            numId: 列表ID
            ilvl: 列表级别
            numFmt: 编号格式
            
        Returns:
            编号字符串
        """
        # 初始化计数器
        if numId not in self.list_counters:
            self.list_counters[numId] = {}
        
        if ilvl not in self.list_counters[numId]:
            self.list_counters[numId][ilvl] = 0
        
        # 如果是编号列表
        if numFmt and numFmt != 'bullet':
            self.list_counters[numId][ilvl] += 1
            
            # 如果是多级列表，计算完整编号（如1.1, 1.2）
            if ilvl > 0:
                parts = []
                for lvl in range(ilvl + 1):
                    if lvl in self.list_counters[numId]:
                        parts.append(str(self.list_counters[numId][lvl]))
                    else:
                        parts.append('1')
                number = '.'.join(parts)
            else:
                number = str(self.list_counters[numId][ilvl])
            
            # 根据numFmt格式化
            if numFmt == 'decimal':
                pass
            elif numFmt == 'lowerLetter':
                number = DocxUtils.number_to_letter(int(number), lower=True)
            elif numFmt == 'upperLetter':
                number = DocxUtils.number_to_letter(int(number), lower=False)
            elif numFmt == 'lowerRoman':
                number = DocxUtils.number_to_roman(int(number), lower=True)
            elif numFmt == 'upperRoman':
                number = DocxUtils.number_to_roman(int(number), lower=False)
            
            return number
        else:
            return None
    
    def _reset_list_counter(self, numId, ilvl):
        """重置指定级别及以下的计数器"""
        if numId in self.list_counters:
            for lvl in list(self.list_counters[numId].keys()):
                if lvl >= ilvl:
                    del self.list_counters[numId][lvl]
    
    def _is_toc_paragraph(self, paragraph):
        """
        检测段落是否为目录项段落
        
        Returns:
            (is_toc, title_text, border_style): 是否为目录段落，提取的标题文本，以及边框样式
        """
        try:
            xml_root = ET.fromstring(paragraph._element.xml)
            
            # 检查是否包含TOC书签
            has_toc_bookmark = False
            for bookmark in xml_root.findall(f'.//{DocxUtils.W_NS}bookmarkStart'):
                name = bookmark.get(f'{DocxUtils.W_NS}name', '')
                if name.startswith('_Toc'):
                    has_toc_bookmark = True
                    self.log_file.write(f"    目录检查: 找到TOC书签 {name}\n")
                    break
            
            # 检查段落样式
            pPr = paragraph._element.pPr
            is_toc_style = False
            if pPr is not None:
                pStyle = pPr.find(qn('w:pStyle'))
                if pStyle is not None:
                    style_val = pStyle.get(qn('w:val'), '')
                    # 检查是否为目录相关样式
                    if 'toc' in style_val.lower() or 'contents' in style_val.lower():
                        is_toc_style = True
                        self.log_file.write(f"    目录检查: 找到目录样式 {style_val}\n")
            
            # 检查段落内容特征（启发式方法）
            raw_text = paragraph.text.strip()
            has_toc_pattern = False
            if raw_text:
                # 匹配目录模式：数字 + . + 标题 + 空格/制表符 + 数字
                import re
                toc_pattern1 = r'^\d+\.\s*[^0-9]+[\t\s]+\d+\s*$'
                toc_pattern2 = r'^\d+\.\s*[^\d]+[\t\s]+\d+$'
                if re.match(toc_pattern1, raw_text) or re.match(toc_pattern2, raw_text):
                    has_toc_pattern = True
                    self.log_file.write(f"    目录检查: 匹配目录模式 '{raw_text}'\n")
            
            # 继续处理边框样式（无论使用哪种检测方式）
            pPr = paragraph._element.pPr
            border_style = ""
            if pPr is not None:
                # 检查段落边框
                pBdr = pPr.find(f'{DocxUtils.W_NS}pBdr')
                if pBdr is not None:
                    border_styles = []
                    self.log_file.write(f"    找到段落边框定义\n")
                    
                    for border_name in ['top', 'left', 'bottom', 'right']:
                        border = pBdr.find(f'{DocxUtils.W_NS}{border_name}')
                        if border is not None:
                            val = border.get(f'{DocxUtils.W_NS}val')
                            self.log_file.write(f"      {border_name}边框: val={val}\n")
                            
                            if val and val != 'nil' and val != 'none':
                                sz = border.get(f'{DocxUtils.W_NS}sz')
                                color = border.get(f'{DocxUtils.W_NS}color')
                                space = border.get(f'{DocxUtils.W_NS}space')
                                
                                border_style_part = f'border-{border_name}:'
                                
                                # 边框宽度
                                if sz:
                                    border_width = int(int(sz) / 8)  # 1/8点 = 1px
                                    if border_width == 0:
                                        border_width = 1
                                    border_style_part += f'{border_width}px'
                                else:
                                    border_style_part += '1px'
                                
                                # 边框样式
                                if val == 'single':
                                    border_style_part += ' solid'
                                elif val == 'double':
                                    border_style_part += ' double'
                                elif val == 'dashed':
                                    border_style_part += ' dashed'
                                elif val == 'dotted':
                                    border_style_part += ' dotted'
                                else:
                                    border_style_part += ' solid'
                                
                                # 边框颜色
                                if color and color != 'auto' and color != 'nil':
                                    if len(color) == 6:  # 十六进制颜色
                                        border_style_part += f' #{color}'
                                    else:
                                        border_style_part += ' #000000'
                                else:
                                    border_style_part += ' #000000'
                                
                                border_styles.append(border_style_part)
                                
                                # 边框间距
                                if space:
                                    space_px = int(int(space) * 0.75)  # 点转像素近似
                                    if border_name == 'top':
                                        border_styles.append(f'padding-top: {space_px}px')
                                    elif border_name == 'bottom':
                                        border_styles.append(f'padding-bottom: {space_px}px')
                                    elif border_name == 'left':
                                        border_styles.append(f'padding-left: {space_px}px')
                                    elif border_name == 'right':
                                        border_styles.append(f'padding-right: {space_px}px')
                    
                    if border_styles:
                        border_style = '; '.join(border_styles)
                        # 添加一些默认的内边距确保边框效果明显
                        if 'padding' not in border_style:
                            border_style += '; padding: 4px 8px'
                        self.log_file.write(f"    目录检查: 找到边框样式 {border_style}\n")
            
            # 检查段落样式，看是否为目录样式
            pPr = paragraph._element.pPr
            border_style = ""
            if pPr is not None:
                # 检查段落边框
                pBdr = pPr.find(f'{DocxUtils.W_NS}pBdr')
                if pBdr is not None:
                    border_styles = []
                    self.log_file.write(f"    找到段落边框定义\n")
                    
                    for border_name in ['top', 'left', 'bottom', 'right']:
                        border = pBdr.find(f'{DocxUtils.W_NS}{border_name}')
                        if border is not None:
                            val = border.get(f'{DocxUtils.W_NS}val')
                            self.log_file.write(f"      {border_name}边框: val={val}\n")
                            
                            if val and val != 'nil' and val != 'none':
                                sz = border.get(f'{DocxUtils.W_NS}sz')
                                color = border.get(f'{DocxUtils.W_NS}color')
                                space = border.get(f'{DocxUtils.W_NS}space')
                                
                                border_style_part = f'border-{border_name}:'
                                
                                # 边框宽度
                                if sz:
                                    border_width = int(int(sz) / 8)  # 1/8点 = 1px
                                    if border_width == 0:
                                        border_width = 1
                                    border_style_part += f'{border_width}px'
                                else:
                                    border_style_part += '1px'
                                
                                # 边框样式
                                if val == 'single':
                                    border_style_part += ' solid'
                                elif val == 'double':
                                    border_style_part += ' double'
                                elif val == 'dashed':
                                    border_style_part += ' dashed'
                                elif val == 'dotted':
                                    border_style_part += ' dotted'
                                else:
                                    border_style_part += ' solid'
                                
                                # 边框颜色
                                if color and color != 'auto' and color != 'nil':
                                    if len(color) == 6:  # 十六进制颜色
                                        border_style_part += f' #{color}'
                                    else:
                                        border_style_part += ' #000000'
                                else:
                                    border_style_part += ' #000000'
                                
                                border_styles.append(border_style_part)
                                
                                # 边框间距
                                if space:
                                    space_px = int(int(space) * 0.75)  # 点转像素近似
                                    if border_name == 'top':
                                        border_styles.append(f'padding-top: {space_px}px')
                                    elif border_name == 'bottom':
                                        border_styles.append(f'padding-bottom: {space_px}px')
                                    elif border_name == 'left':
                                        border_styles.append(f'padding-left: {space_px}px')
                                    elif border_name == 'right':
                                        border_styles.append(f'padding-right: {space_px}px')
                    
                    if border_styles:
                        border_style = '; '.join(border_styles)
                        # 添加一些默认的内边距确保边框效果明显
                        if 'padding' not in border_style:
                            border_style += '; padding: 4px 8px'
                        self.log_file.write(f"    检查边框样式: {border_style}\n")
            
            if (has_toc_bookmark and (is_toc_style or has_toc_pattern)) or (is_toc_style and has_toc_pattern) or (has_toc_pattern and ('\t' in raw_text or '  ' in raw_text)):
                # 获取段落的完整文本内容
                raw_text = paragraph.text.strip()
                self.log_file.write(f"    目录原始文本: '{raw_text}'\n")
                
                # 移除制表符（&emsp;&emsp;实际上是制表符或空格）
                # 先处理HTML实体编码的制表符
                cleaned_text = raw_text.replace('&emsp;', '').replace('&ensp;', '').replace('\u2003', '')
                cleaned_text = re.sub(r'[\t ]{2,}', '', cleaned_text)
                cleaned_text = re.sub(r'^\s+', '', cleaned_text)
                
                # 移除页码（包括点号分隔的页码，以及HTML实体形式的空格）
                # 匹配模式：连续的点号后跟数字，或者只有数字，包括&nbsp;
                title_text = re.sub(r'[.\u2026]+(?:&nbsp;|\u00A0)*\d+\s*$', '', cleaned_text)
                title_text = re.sub(r'(?:&nbsp;|\u00A0|\s)+\d+\s*$', '', title_text)
                
                # 移除末尾的点号和制表符
                title_text = re.sub(r'[.\u2026\u2003]+$', '', title_text)
                title_text = title_text.strip()
                
                # 移除中间的多余空格、制表符和HTML实体
                title_text = re.sub(r'(?:&nbsp;|\u00A0|\s)+', '', title_text)
                title_text = title_text.strip()
                
                self.log_file.write(f"    目录清理后文本: '{title_text}'\n")
                
                if title_text and len(title_text) > 1:
                    return True, title_text, border_style
        
        except Exception as e:
            self.log_file.write(f"    目录检查异常: {e}\n")
        
        return False, None, None
    
    def _should_merge_paragraphs(self, p1, p2):
        """判断两个段落是否应该合并（用于处理连续的浮动图片/文本框）"""
        try:
            # 检查p1是否包含浮动drawing
            has_float1 = False
            for drawing in p1.findall(f'.//{DocxUtils.W_NS}drawing'):
                if drawing.find(f'.//{DocxUtils.WP_NS}anchor') is not None:
                    has_float1 = True
                    break
            
            if not has_float1:
                return False
                
            # 检查p2是否包含浮动drawing
            has_float2 = False
            for drawing in p2.findall(f'.//{DocxUtils.W_NS}drawing'):
                if drawing.find(f'.//{DocxUtils.WP_NS}anchor') is not None:
                    has_float2 = True
                    break
                    
            if not has_float2:
                return False
            
            # 检查p2是否有直接的文本内容（不包含drawing内的文本）
            # 如果p2有直接文本（如标题、说明文字），则不应合并，应保留为独立段落
            has_direct_text = False
            for r in p2.findall(f'{DocxUtils.W_NS}r'):
                for t in r.findall(f'{DocxUtils.W_NS}t'):
                    if t.text and t.text.strip():
                        has_direct_text = True
                        break
                if has_direct_text:
                    break
            
            if has_direct_text:
                return False
            
            return True
        except Exception:
            return False

    def _merge_paragraph_elements(self, target_p, source_p):
        """将source_p的内容合并到target_p"""
        for run in source_p.findall(f'{DocxUtils.W_NS}r'):
            target_p.append(run)

    def _traverse_document_body(self, body_element, table_context=None, progress=None):
        """
        递归遍历文档body元素，按文档实际顺序处理所有元素
        
        Args:
            body_element: 文档body元素（任意XML元素）
            table_context: 表格上下文信息，用于表格嵌套情况
            
        Returns:
            HTML内容列表
        """
        html_parts = []
        element_count = 0
        
        self.log_file.write(f"\n--- 开始遍历文档元素，共 {len(body_element)} 个子元素 ---\n")
        
        # 使用索引遍历，以便跳过已合并的元素
        i = 0
        total_len = len(body_element)
        
        while i < total_len:
            elem = body_element[i]
            element_count += 1
            self.log_file.write(f"处理第 {element_count} 个元素: {elem.tag}\n")
            if progress is not None:
                try:
                    progress.update(1)
                except Exception:
                    pass
            
            # 处理段落 <w:p>
            if elem.tag == DocxUtils.W_NS + 'p':
                self.log_file.write(f"  -> 发现段落元素\n")
                
                # 尝试合并后续连续的浮动段落
                next_idx = i + 1
                while next_idx < total_len:
                    next_elem = body_element[next_idx]
                    if next_elem.tag == DocxUtils.W_NS + 'p' and self._should_merge_paragraphs(elem, next_elem):
                        self.log_file.write(f"  -> 合并后续浮动段落: {next_idx}\n")
                        self._merge_paragraph_elements(elem, next_elem)
                        # 更新进度条，因为我们跳过了一个元素
                        if progress is not None:
                            try:
                                progress.update(1)
                            except Exception:
                                pass
                        next_idx += 1
                    else:
                        break
                
                # 更新主循环索引
                i = next_idx
                
                from docx.text.paragraph import Paragraph
                para = Paragraph(elem, self.doc)
                html_parts.append(self.process_paragraph(para, table_context=table_context))
            
            # 处理表格 <w:tbl>
            elif elem.tag == DocxUtils.W_NS + 'tbl':
                self.log_file.write(f"  -> 发现表格元素\n")
                from docx.table import Table
                tbl = Table(elem, self.doc)
                html_parts.append(self.process_table(tbl, len(html_parts) + 1, table_context=table_context))
                i += 1
            
            # 处理文本框 <w:txbxContent> - 递归处理
            elif elem.tag == DocxUtils.W_NS + 'txbxContent':
                self.log_file.write(f"  -> 发现文本框元素，递归处理\n")
                # 递归处理文本框内容
                nested_html = self._traverse_document_body(elem, table_context)
                if nested_html:
                    style = self._get_txbx_div_style(elem)
                    if style:
                        html_parts.append(f'<div style="{style}">')
                    else:
                        html_parts.append('<div>')
                    html_parts.extend(nested_html)
                    html_parts.append('</div>')
                i += 1
            
            # 处理其他自定义XML元素，递归处理其子元素
            elif len(elem) > 0:
                self.log_file.write(f"  -> 发现复合元素，递归处理 {len(elem)} 个子元素\n")
                # 递归处理子元素
                nested_html = self._traverse_document_body(elem, table_context)
                if nested_html:
                    html_parts.extend(nested_html)
                i += 1
            
            # 跳过其他叶子元素
            else:
                self.log_file.write(f"  -> 跳过叶子元素: {elem.tag}\n")
                i += 1
        
        self.log_file.write(f"--- 遍历完成，共处理 {element_count} 个元素，生成 {len(html_parts)} 个HTML片段 ---\n")
        return html_parts
    
    def extract_paragraph_text_with_links(self, paragraph):
        """
        提取段落文本，支持换行符、制表符、符号等
        保持原始文本顺序和格式，避免序号被分割
        
        Returns:
            HTML字符串
        """
        html_parts = []
        
        for run_idx, run in enumerate(paragraph.runs):
            # 获取run的样式
            style = self._get_run_format_from_xml(run)
            
            # 检查是否包含超链接
            hyperlink_elem = run._element.find(f'.//{DocxUtils.W_NS}hyperlink')
            is_hyperlink = hyperlink_elem is not None
            
            # 遍历run的所有子元素，分别处理文本和其他元素
            text_buffer = []
            self._process_run_elements(run._element, text_buffer, is_hyperlink, style, html_parts)
            
            # 处理缓冲区中剩余的文本
            if text_buffer:
                combined_text = ''.join(text_buffer)
                if combined_text.strip():
                    if is_hyperlink:
                        self._process_hyperlink(hyperlink_elem, combined_text, style, html_parts)
                    else:
                        escaped_text = escape(combined_text)
                        html_parts.append(f'<span style="{style}">{escaped_text}</span>')
        
        return "".join(html_parts)
    
    def _process_run_elements(self, run_element, text_buffer, is_hyperlink, style, html_parts):
        """处理run内的所有元素，区分文本和格式元素"""
        for child in run_element:
            # 文本元素
            if child.tag == DocxUtils.W_NS + 't':
                text = child.text or ''
                # 检查父级是否为超链接
                hyperlink_parent = self._find_parent_with_tag(child, [DocxUtils.W_NS + 'hyperlink'])
                if hyperlink_parent is not None:
                    # 将缓冲区先输出
                    if text_buffer:
                        combined_text = ''.join(text_buffer)
                        if combined_text.strip():
                            escaped_text = escape(combined_text)
                            html_parts.append(f'<span style="{style}">{escaped_text}</span>')
                        text_buffer.clear()
                    # 输出超链接文本
                    rId = hyperlink_parent.get(f'{DocxUtils.R_NS}id')
                    anchor = hyperlink_parent.get(f'{DocxUtils.W_NS}anchor')
                    link_text = text
                    if link_text.strip():
                        fake_run_style = style
                        self._process_hyperlink(hyperlink_parent, link_text, fake_run_style, html_parts)
                    continue
                # 检查文本是否为点号或句号
                if text in ['.', '。', '、', '，']:
                    text_buffer.append(text)
                    text_buffer.append('  ')
                else:
                    text_buffer.append(text)
            
            # 换行符 - 立即添加真实的HTML换行标签，不放入文本缓冲区
            elif child.tag == DocxUtils.W_NS + 'br':
                # 先处理缓冲区中的文本
                if text_buffer:
                    combined_text = ''.join(text_buffer)
                    if combined_text.strip():
                        if is_hyperlink:
                            self._process_hyperlink(None, combined_text, style, html_parts)
                        else:
                            escaped_text = escape(combined_text)
                            html_parts.append(f'<span style="{style}">{escaped_text}</span>')
                    text_buffer.clear()
                
                # 添加换行标签（不带任何额外内容）
                html_parts.append('<br>')
            
            # 制表符 - 转换为HTML实体，但先处理缓冲区文本
            elif child.tag == DocxUtils.W_NS + 'tab':
                # 先处理缓冲区中的文本
                if text_buffer:
                    combined_text = ''.join(text_buffer)
                    if combined_text.strip():
                        if is_hyperlink:
                            self._process_hyperlink(None, combined_text, style, html_parts)
                        else:
                            escaped_text = escape(combined_text)
                            html_parts.append(f'<span style="{style}">{escaped_text}</span>')
                    text_buffer.clear()
                
                # 添加制表符（使用CSS方式实现）
                html_parts.append(f'<span style="{style}">&emsp;&emsp;</span>')
            
            # 符号 - 立即处理
            elif child.tag == DocxUtils.W_NS + 'sym':
                font = child.get(f'{DocxUtils.W_NS}font')
                char = child.get(f'{DocxUtils.W_NS}char')
                if char:
                    # 先处理缓冲区中的文本
                    if text_buffer:
                        combined_text = ''.join(text_buffer)
                        if combined_text.strip():
                            if is_hyperlink:
                                self._process_hyperlink(None, combined_text, style, html_parts)
                            else:
                                escaped_text = escape(combined_text)
                                html_parts.append(f'<span style="{style}">{escaped_text}</span>')
                        text_buffer.clear()
                    
                    # 添加符号
                    html_parts.append(f'<span style="{style}">&#x{char};</span>')
            
            # 字段分隔符 - 忽略
            elif child.tag == DocxUtils.W_NS + 'fldChar':
                pass
    
    def _process_hyperlink(self, hyperlink_elem, text, style, html_parts):
        """处理超链接"""
        if not hyperlink_elem:
            escaped_text = escape(text)
            html_parts.append(f'<span style="{style}">{escaped_text}</span>')
            return
            
        rId = hyperlink_elem.get(f'{DocxUtils.R_NS}id')
        anchor = hyperlink_elem.get(f'{DocxUtils.W_NS}anchor')
        
        if rId:
            for rel in self.doc.part.rels.values():
                if rel.rId == rId:
                    hyperlink = rel.target_ref
                    
                    # 处理不同类型的链接
                    href = None
                    if anchor:
                        href = f"#{anchor}"
                        self.log_file.write(f"  书签链接: #{anchor}\n")
                    elif hyperlink and hyperlink.startswith('mailto:'):
                        href = hyperlink
                        self.log_file.write(f"  邮件链接: {href}\n")
                    elif hyperlink:
                        href = hyperlink
                        self.log_file.write(f"  普通链接: {href}\n")
                    
                    if href:
                        escaped_text = escape(text)
                        escaped_href = escape(href)
                        html_parts.append(f'<a href="{escaped_href}" style="{style}">{escaped_text}</a>')
                    else:
                        escaped_text = escape(text)
                        html_parts.append(f'<span style="{style}">{escaped_text}</span>')
                    break
        else:
            escaped_text = escape(text)
            html_parts.append(f'<span style="{style}">{escaped_text}</span>')
    
    def _get_run_format_from_xml(self, run):
        """从XML中获取run格式（合并直接格式和样式）"""
        rPr = run._element.rPr
        css_parts = []
        
        # 获取run样式
        style_name = None
        if run.style:
            style_name = run.style.name
        
        # 从样式定义中获取基础格式（使用继承链）
        if style_name and style_name in self.style_inheritance:
            inherited_css = self._get_resolved_style_css(style_name, 'run')
            if inherited_css:
                css_parts.append(inherited_css)
        
        # 获取run上的直接格式（优先级最高）
        if rPr is not None:
            direct_css = DocxUtils.get_run_style_css(rPr, apply_scale=False)
            if direct_css:
                css_parts.append(direct_css)
        
        # 合并所有CSS
        if css_parts:
            return DocxUtils.merge_css_with_priority(css_parts)
        
        return ""
    
    def _get_paragraph_format_from_xml(self, paragraph):
        """
        获取段落格式（合并直接格式和样式）
        
        Returns:
            CSS字符串
        """
        pPr = paragraph._element.pPr
        css_parts = []
        
        # 获取段落样式
        style_name = None
        if paragraph.style:
            style_name = paragraph.style.name
        
        # 从样式定义中获取基础格式（使用继承链）
        if style_name and style_name in self.style_inheritance:
            inherited_css = self._get_resolved_style_css(style_name, 'paragraph')
            if inherited_css:
                css_parts.append(inherited_css)
        
        # 获取段落上的直接格式（优先级最高）
        if pPr is not None:
            direct_css = DocxUtils.get_paragraph_style_css(pPr, self.log_file, apply_scale=False)
            if direct_css:
                css_parts.append(direct_css)
        
        # 合并所有CSS
        if css_parts:
            return DocxUtils.merge_css_with_priority(css_parts)

        return ""
    
    def is_heading_from_xml(self, paragraph):
        """从XML判断段落是否为标题"""
        pPr = paragraph._element.pPr
        
        if pPr is None:
            return False
        
        pStyle = pPr.find(qn('w:pStyle'))
        if pStyle is not None:
            style_val = pStyle.get(qn('w:val'), '')
            self.log_file.write(f"    标题检查: 找到样式 {style_val}\n")
            if style_val and ('heading' in style_val.lower() or 
                           'title' in style_val.lower() or
                           style_val.lower().startswith('head')):
                self.log_file.write(f"    标题检查: 样式匹配 {style_val}\n")
                return True
        
        outlineLvl = pPr.find(qn('w:outlineLvl'))
        if outlineLvl is not None:
            lvl_val = outlineLvl.get(qn('w:val'))
            if lvl_val is not None:
                self.log_file.write(f"    标题检查: 大纲级别 {lvl_val}\n")
                return True
        
        is_large_font = False
        is_bold = False
        
        for run in paragraph.runs:
            rPr = run._element.rPr
            if rPr is not None:
                sz = rPr.find(qn('w:sz'))
                if sz is not None:
                    sz_val = sz.get(qn('w:val'))
                    if sz_val:
                        font_size_pt = int(sz_val) / 2
                        if font_size_pt >= 15:
                            is_large_font = True
                            self.log_file.write(f"    标题检查: 发现大字体 {font_size_pt}pt\n")
                
                b = rPr.find(qn('w:b'))
                if b is not None and b.get(qn('w:val')) != '0':
                    is_bold = True
                    self.log_file.write(f"    标题检查: 发现加粗\n")
                
                if is_large_font and is_bold:
                    self.log_file.write(f"    标题检查: 大字体+加粗，判定为标题\n")
                    return True
        
        try:
            xml_root = ET.fromstring(paragraph._element.xml)
            has_toc_bookmark = False
            bookmark_name = None
            for bookmark in xml_root.findall(f'.//{DocxUtils.W_NS}bookmarkStart'):
                name = bookmark.get(f'{DocxUtils.W_NS}name', '')
                if name and name.startswith('_Toc'):
                    has_toc_bookmark = True
                    bookmark_name = name
                    break
            if has_toc_bookmark:
                raw_text = paragraph.text.strip()
                is_toc_like = False
                if raw_text:
                    toc_pattern1 = r'^\d+\.\s*[^0-9]+[\t\s]+\d+\s*$'
                    toc_pattern2 = r'^\d+\.\s*[^\d]+[\t\s]+\d+$'
                    if re.match(toc_pattern1, raw_text) or re.match(toc_pattern2, raw_text):
                        is_toc_like = True
                if not is_toc_like:
                    self.log_file.write(f"    标题检查: 段落包含TOC书签 {bookmark_name}\n")
                    return True
        except Exception:
            pass
        
        return False
    
    def _generate_heading_id(self, heading_html):
        """
        为标题生成唯一ID
        
        Args:
            heading_html: 标题的HTML文本
            
        Returns:
            唯一的ID字符串
        """
        # 提取纯文本作为基础
        heading_text = DocxUtils.strip_html_tags(heading_html).strip()
        
        # 移除特殊字符
        safe_text = re.sub(r'[^\w\u4e00-\u9fff-]', '_', heading_text)
        safe_text = re.sub(r'_+', '_', safe_text).strip('_')
        
        # 如果为空，使用默认值
        if not safe_text:
            safe_text = 'heading'
        
        # 生成唯一ID
        base_id = safe_text
        counter = 1
        heading_id = base_id
        
        while heading_id in self.headings_map.values():
            heading_id = f"{base_id}_{counter}"
            counter += 1
        
        return heading_id
    
    def _find_heading_id(self, title_text):
        """
        根据目录条目文本查找对应的标题ID
        
        Args:
            title_text: 目录条目的标题文本
            
        Returns:
            匹配的标题ID，如果未找到则返回None
        """
        # 首先尝试精确匹配
        if title_text in self.headings_map:
            return self.headings_map[title_text]
            
        # 尝试忽略空格的精确匹配
        title_norm = title_text.replace(' ', '').replace('\t', '')
        for heading_text, heading_id in self.headings_map.items():
            if heading_text.replace(' ', '').replace('\t', '') == title_norm:
                return heading_id
        
        # 尝试包含匹配
        for heading_text, heading_id in self.headings_map.items():
            if title_text.strip() in heading_text.strip() or heading_text.strip() in title_text.strip():
                return heading_id
        
        # 最后尝试关键词匹配
        title_keywords = re.findall(r'[\u4e00-\u9fff]+|[a-zA-Z]+', title_text)
        if title_keywords:
            for heading_text, heading_id in self.headings_map.items():
                heading_keywords = re.findall(r'[\u4e00-\u9fff]+|[a-zA-Z]+', heading_text)
                
                if heading_keywords:
                    common_keywords = set(title_keywords) & set(heading_keywords)
                    if len(common_keywords) >= 2 or (len(common_keywords) >= 1 and len(title_keywords) <= 2):
                        return heading_id
        
        return None
    
    def parse_list_from_xml(self, paragraph):
        """
        从XML中解析列表信息
        
        Returns:
            tuple: (list_type, level, list_id, list_color, numFmt, bullet_size)
            - numbered: bullet_size为None，numFmt为编号格式
            - bulleted: numFmt为None，bullet_size为项目符号字号（pt）
        """
        pPr = paragraph._element.pPr
        if pPr is None:
            return None, None, None, None, None, None
        
        # 检查编号列表（段落自身）
        numPr = pPr.find(qn('w:numPr'))
        if numPr is not None:
            ilvl_elem = numPr.find(qn('w:ilvl'))
            ilvl = int(ilvl_elem.get(qn('w:val'))) if ilvl_elem is not None else 0
            
            numId_elem = numPr.find(qn('w:numId'))
            numId = int(numId_elem.get(qn('w:val'))) if numId_elem is not None else 0
            
            # 从numbering定义中判断
            if numId in self.numbering_definitions:
                levels = self.numbering_definitions[numId]
                if ilvl in levels:
                    level_info = levels[ilvl]
                    is_bullet = level_info['is_bullet']
                    bullet_char = level_info['bullet_char']
                    numFmt = level_info.get('numFmt')
                    bullet_color = level_info.get('bullet_color')
                    bullet_size = level_info.get('bullet_size')
                    
                    paragraph_color = self._get_list_number_color_from_xml(paragraph)
                    
                    if is_bullet:
                        return 'bulleted', ilvl, bullet_char, bullet_color or paragraph_color, None, bullet_size
                    else:
                        return 'numbered', ilvl, numId, paragraph_color, numFmt, None
            
            # 默认为编号列表
            list_color = self._get_list_number_color_from_xml(paragraph)
            return 'numbered', ilvl, numId, list_color, None, None
        
        # 若段落未显式设置numPr，尝试从样式继承链获取
        try:
            style_val = None
            if pPr is not None:
                pStyle = pPr.find(qn('w:pStyle'))
                if pStyle is not None:
                    style_val = pStyle.get(qn('w:val'), '')
            if style_val and style_val in self.style_inheritance:
                chain = self.style_inheritance.get(style_val, [])
                for style_id in chain:
                    style_def = self.style_definitions.get(style_id)
                    if not style_def:
                        continue
                    style_pPr = style_def.get('pPr')
                    if style_pPr is not None:
                        style_numPr = style_pPr.find(qn('w:numPr'))
                        if style_numPr is not None:
                            ilvl_elem = style_numPr.find(qn('w:ilvl'))
                            ilvl = int(ilvl_elem.get(qn('w:val'))) if ilvl_elem is not None else 0
                            
                            numId_elem = style_numPr.find(qn('w:numId'))
                            numId = int(numId_elem.get(qn('w:val'))) if numId_elem is not None else 0
                            
                            if numId in self.numbering_definitions:
                                levels = self.numbering_definitions[numId]
                                if ilvl in levels:
                                    level_info = levels[ilvl]
                                    is_bullet = level_info['is_bullet']
                                    bullet_char = level_info['bullet_char']
                                    numFmt = level_info.get('numFmt')
                                    bullet_color = level_info.get('bullet_color')
                                    bullet_size = level_info.get('bullet_size')
                                    
                                    paragraph_color = self._get_list_number_color_from_xml(paragraph)
                                    
                                    if is_bullet:
                                        return 'bulleted', ilvl, bullet_char, bullet_color or paragraph_color, None, bullet_size
                                    else:
                                        return 'numbered', ilvl, numId, paragraph_color, numFmt, None
                            # 找到样式级numPr但未能映射定义，直接返回编号默认值
                            list_color = self._get_list_number_color_from_xml(paragraph)
                            return 'numbered', ilvl, numId, list_color, None, None
        except Exception:
            pass
        
        # 检查项目符号
        buChar = pPr.find(qn('w:buChar'))
        if buChar is not None:
            bullet_char = buChar.get(qn('w:char'), '•')
            ilvl_elem = pPr.find(qn('w:ilvl'))
            ilvl = int(ilvl_elem.get(qn('w:val'))) if ilvl_elem is not None else 0
            list_color = self._get_list_number_color_from_xml(paragraph)
            return 'bulleted', ilvl, bullet_char, list_color, None, None
        
        return None, None, None, None, None, None
    
    def _matches_toc_title(self, text_plain):
        """判断纯文本是否匹配已记录的目录标题（宽松匹配）"""
        try:
            if not text_plain:
                return None
            # 规范化：移除所有空白与标点，仅保留中英文与数字
            def norm(s):
                s = s.strip()
                s = re.sub(r'\s+', '', s)
                s = re.sub(r'[^\w\u4e00-\u9fff]', '', s)
                return s
            norm_text = norm(text_plain)
            for t in list(self.toc_titles):
                nt = norm(t)
                if nt and (norm_text == nt or norm_text in nt or nt in norm_text):
                    return t
        except Exception:
            pass
        return None

    def _paragraph_has_hyperlink(self, paragraph):
        """判断段落是否包含原始超链接元素"""
        try:
            elem = paragraph._element
            if elem is None:
                return False
            # 检查是否存在w:hyperlink或HYPERLINK字段
            has_link_elem = elem.find(f'.//{DocxUtils.W_NS}hyperlink') is not None
            if has_link_elem:
                return True
            # 检查字段指令（fldSimple 或 instrText 包含 HYPERLINK）
            if elem.find(f'.//{DocxUtils.W_NS}fldSimple') is not None:
                fld = elem.find(f'.//{DocxUtils.W_NS}fldSimple')
                instr = fld.get(f'{DocxUtils.W_NS}instr')
                if instr and 'HYPERLINK' in instr:
                    return True
            for instrText in elem.findall(f'.//{DocxUtils.W_NS}instrText'):
                if instrText.text and 'HYPERLINK' in instrText.text:
                    return True
        except Exception:
            pass
        return False

    def _get_list_number_color_from_xml(self, paragraph):
        """从XML中获取列表序号的颜色"""
        if not paragraph.runs:
            return None
        
        first_run = paragraph.runs[0]
        rPr = first_run._element.rPr
        
        if rPr is not None:
            color = rPr.find(qn('w:color'))
            if color is not None:
                color_val = color.get(qn('w:val'))
                if color_val and len(color_val) == 6:
                    return f"#{color_val}"
        
        return None
    
    def extract_image_with_textbox(self, r_embed, drawing_elem, image_index):
        """
        提取图片及其显示尺寸，并解析叠加的文本框位置（不再进行合成）
        
        Args:
            r_embed: 图片关系ID
            drawing_elem: drawing元素
            image_index: 图片索引
            
        Returns:
            dict: {filename, width, height, overlays: [{'text': str, 'x': px, 'y': px, 'width': px|None, 'height': px|None}]}
        """
        try:
            image_part = None
            for rel in self.doc.part.rels.values():
                if rel.rId == r_embed:
                    image_part = rel.target_part
                    break
            if image_part is None:
                self.log_file.write(f"      未找到图片关系: {r_embed}\n")
                return None, None, None
            
            image_bytes = image_part.blob
            img = PILImage.open(BytesIO(image_bytes))
            original_width, original_height = img.size
            
            crop_rect = self._get_image_crop(drawing_elem)
            if crop_rect:
                try:
                    l, t, r, b = crop_rect
                    box = (
                        int(l * original_width),
                        int(t * original_height),
                        int(r * original_width),
                        int(b * original_height),
                    )
                    img = img.crop(box)
                    original_width, original_height = img.size
                except Exception:
                    pass
            
            rotation = self._get_image_rotation(drawing_elem)
            try:
                if rotation and abs(rotation) > 0.01:
                    img = img.rotate(rotation, expand=True)
                    original_width, original_height = img.size
            except Exception:
                pass
            
            display_width, display_height = self._get_image_display_size(
                drawing_elem, original_width, original_height
            )
            
            textboxes, positions = self.extract_textbox_from_drawing(
                drawing_elem, display_width, display_height
            )
            filename = f"image_{image_index}.png"
            save_path = os.path.join(self.images_dir, filename)
            try:
                img.save(save_path)
            except Exception:
                filename = f"image_{image_index}.jpg"
                save_path = os.path.join(self.images_dir, filename)
                img.convert("RGB").save(save_path, quality=90)
            
            overlays = []
            if textboxes and positions:
                for tb, pos in zip(textboxes, positions):
                    overlay = {
                        'type': 'text',
                        'html': tb.get('html', ''),
                        'x': int(pos.get('x', 0)),
                        'y': int(pos.get('y', 0)),
                        'width': int(pos.get('width')) if pos.get('width') else None,
                        'height': int(pos.get('height')) if pos.get('height') else None
                    }
                    if tb.get('border_css'):
                        overlay['border_css'] = tb.get('border_css')
                    elif hasattr(self, '_get_shape_border_css'):
                        # 兼容旧代码，如果有定义该方法
                         border_css = self._get_shape_border_css(tb.get('elem'))
                         if border_css:
                            overlay['border_css'] = border_css
                    overlays.append(overlay)
            
            # 提取同一drawing中的其他图片作为叠加层（相对于主图片定位）
            try:
                original_img_width, original_img_height = self._get_original_image_size(drawing_elem)
                scale_x = display_width / original_img_width if original_img_width > 0 else 1
                scale_y = display_height / original_img_height if original_img_height > 0 else 1
                overlay_index = 0
                for elem in drawing_elem.iter():
                    if elem.tag == DocxUtils.PIC_NS + 'pic':
                        blip = elem.find(f'.//{DocxUtils.A_NS}blip')
                        if blip is None:
                            continue
                        o_embed = blip.get(f'{DocxUtils.R_NS}embed')
                        if not o_embed or o_embed == r_embed:
                            continue
                        overlay_part = None
                        for rel in self.doc.part.rels.values():
                            if rel.rId == o_embed:
                                overlay_part = rel.target_part
                                break
                        if overlay_part is None:
                            continue
                        overlay_bytes = overlay_part.blob
                        try:
                            oimg = PILImage.open(BytesIO(overlay_bytes))
                        except Exception:
                            continue
                        # 读取位置与大小（未缩放）
                        pos_info = self._get_textbox_position(elem, display_width, display_height, drawing_elem)
                        ox = int(pos_info.get('x', 0))
                        oy = int(pos_info.get('y', 0))
                        owidth = int(pos_info.get('width')) if pos_info.get('width') else None
                        oheight = int(pos_info.get('height')) if pos_info.get('height') else None
                        # 保存叠加图片
                        overlay_index += 1
                        ofilename = f"image_{image_index}_overlay_{overlay_index}.png"
                        os_path = os.path.join(self.images_dir, ofilename)
                        try:
                            oimg.save(os_path)
                        except Exception:
                            ofilename = f"image_{image_index}_overlay_{overlay_index}.jpg"
                            os_path = os.path.join(self.images_dir, ofilename)
                            try:
                                oimg.convert("RGB").save(os_path, quality=90)
                            except Exception:
                                continue
                        overlays.append({
                            'type': 'image',
                            'filename': ofilename,
                            'x': ox,
                            'y': oy,
                            'width': owidth,
                            'height': oheight
                        })
            except Exception as e:
                self.log_file.write(f"      提取叠加图片失败: {e}\n")
            return {
                'filename': filename,
                'width': display_width,
                'height': display_height,
                'overlays': overlays
            }
        except Exception as e:
            self.log_file.write(f"      提取图片失败: {e}\n")
            import traceback
            self.log_file.write(f"      异常详情: {traceback.format_exc()}\n")
            return None
    
    def _get_image_display_size(self, drawing_elem, original_width, original_height):
        """
        获取图片在文档中的显示尺寸
        
        Args:
            drawing_elem: drawing元素
            original_width: 图片原始宽度
            original_height: 图片原始高度
            
        Returns:
            tuple: (display_width, display_height) 显示尺寸（像素）
        """
        try:
            # 查找extent元素（wp:extent）
            extent = drawing_elem.find(f'.//{DocxUtils.WP_NS}extent')
            if extent is not None:
                cx = extent.get('cx')
                cy = extent.get('cy')
                
                if cx and cy:
                    # EMU转像素（图片需要随内容区缩放）
                    width_emu = int(cx)
                    height_emu = int(cy)
                    display_width = DocxUtils.emu_to_pixels(width_emu, apply_scale=True)
                    display_height = DocxUtils.emu_to_pixels(height_emu, apply_scale=True)
                    self.log_file.write(f"      从extent获取显示尺寸: {display_width}x{display_height}\n")
                    return display_width, display_height
            
            # 如果没有extent，尝试从a:ext获取（DrawingML）
            for elem in drawing_elem.iter():
                if elem.tag == DocxUtils.A_NS + 'ext':
                    cx = elem.get('cx')
                    cy = elem.get('cy')
                    if cx and cy:
                        width_emu = int(cx)
                        height_emu = int(cy)
                        display_width = DocxUtils.emu_to_pixels(width_emu, apply_scale=True)
                        display_height = DocxUtils.emu_to_pixels(height_emu, apply_scale=True)
                        self.log_file.write(f"      从a:ext获取显示尺寸: {display_width}x{display_height}\n")
                        return display_width, display_height
            
            # 如果都没有，返回原始尺寸
            self.log_file.write(f"      使用原始尺寸: {original_width}x{original_height}\n")
            return original_width, original_height
            
        except Exception as e:
            self.log_file.write(f"      获取显示尺寸失败: {e}\n")
            return original_width, original_height
    
    def _get_image_rotation(self, drawing_elem):
        """获取图片旋转角度（度）"""
        try:
            # 查找spPr中的a:xfrm元素
            for elem in drawing_elem.iter():
                if elem.tag == DocxUtils.A_NS + 'xfrm':
                    rot = elem.get('rot')
                    if rot:
                        # rot单位是1/60000度
                        return int(rot) / 60000
        except:
            pass
        return 0
    
    def _get_original_image_size(self, drawing_elem):
        """
        获取图片的原始尺寸（未应用缩放的EMU值对应的像素尺寸）
        
        Args:
            drawing_elem: drawing元素
            
        Returns:
            tuple: (original_width, original_height) 原始像素尺寸
        """
        try:
            # 查找extent元素（wp:extent）
            extent = drawing_elem.find(f'.//{DocxUtils.WP_NS}extent')
            if extent is not None:
                cx = extent.get('cx')
                cy = extent.get('cy')
                
                if cx and cy:
                    # EMU转像素（不应用缩放）
                    width_emu = int(cx)
                    height_emu = int(cy)
                    original_width = DocxUtils.emu_to_pixels(width_emu, apply_scale=False)
                    original_height = DocxUtils.emu_to_pixels(height_emu, apply_scale=False)
                    return original_width, original_height
            
            # 如果没有extent，尝试从a:ext获取（DrawingML）
            for elem in drawing_elem.iter():
                if elem.tag == DocxUtils.A_NS + 'ext':
                    cx = elem.get('cx')
                    cy = elem.get('cy')
                    if cx and cy:
                        width_emu = int(cx)
                        height_emu = int(cy)
                        original_width = DocxUtils.emu_to_pixels(width_emu, apply_scale=False)
                        original_height = DocxUtils.emu_to_pixels(height_emu, apply_scale=False)
                        return original_width, original_height
            
            # 如果都没有，返回默认尺寸
            return 800, 600
        except Exception:
            return 800, 600
    
    def _get_image_crop(self, drawing_elem):
        """获取图片裁剪区域"""
        try:
            # 查找srcRect元素（裁剪矩形）
            for elem in drawing_elem.iter():
                if elem.tag == DocxUtils.A_NS + 'srcRect':
                    # l, t, r, b单位是百分比（1/100000）
                    l = int(elem.get('l', 0)) / 100000
                    t = int(elem.get('t', 0)) / 100000
                    r = int(elem.get('r', 0)) / 100000
                    b = int(elem.get('b', 0)) / 100000
                    
                    # 获取原始尺寸
                    blip = drawing_elem.find(f'.//{DocxUtils.A_NS}blip')
                    if blip is not None:
                        # 这里简化处理，假设裁剪比例
                        # 实际需要获取图片原始尺寸
                        if l > 0 or t > 0 or r > 0 or b > 0:
                            # 返回裁剪矩形（简化版）
                            return (l, t, 1-r, 1-b)
        except:
            pass
        return None
    
    def extract_textbox_from_drawing(self, drawing_elem, img_width, img_height):
        """
        从drawing元素中提取文本框内容和位置
        
        Args:
            drawing_elem: drawing元素
            img_width: 图片宽度
            img_height: 图片高度
            
        Returns:
            tuple: (textboxes, positions)
        """
        textboxes = []
        positions = []
        
        try:
            self.log_file.write(f"    开始提取文本框... 图片尺寸: {img_width}x{img_height}\n")
            
            # 方法1: 查找drawing元素内的所有嵌套结构，包括文本框
            # 先查找inline或anchor内的文本框
            for elem in drawing_elem.iter():
                # 查找w:txbxContent（文本框内容）
                if elem.tag == DocxUtils.W_NS + 'txbxContent':
                    self.log_file.write(f"      找到txbxContent元素\n")
                    inner_html_parts = self._traverse_document_body(elem)
                    inner_html = "".join(inner_html_parts).strip()
                    position = self._get_textbox_position(elem, img_width, img_height, drawing_elem)
                    positions.append(position)
                    textboxes.append({
                        'elem': elem,
                        'html': inner_html
                    })
                    self.log_file.write(f"    提取文本框HTML长度: {len(inner_html)}, 位置: {position}\n")
                
                # 方法2: 查找w:txbx（文本框）
                elif elem.tag == DocxUtils.W_NS + 'txbx':
                    txbx_content = elem.find(f'.//{DocxUtils.W_NS}txbxContent')
                    if txbx_content is not None:
                        self.log_file.write(f"      找到txbx元素，包含txbxContent\n")
                        inner_html_parts = self._traverse_document_body(txbx_content)
                        inner_html = "".join(inner_html_parts).strip()
                        position = self._get_textbox_position(elem, img_width, img_height, drawing_elem)
                        positions.append(position)
                        textboxes.append({
                            'elem': elem,
                            'html': inner_html
                        })
                        self.log_file.write(f"    提取文本框HTML长度: {len(inner_html)}, 位置: {position}\n")
                
                # 方法3: 查找wp:docPr（文档属性）可能包含图片标注
                elif elem.tag == DocxUtils.WP_NS + 'docPr':
                    name = elem.get('name', '') or ''
                    description = elem.get('description', '') or ''
                    text_content = name or description
                    if text_content:
                        import re
                        if re.match(r'^\s*(图片|图|Image|Figure)\s*\d+\s*$', text_content, re.IGNORECASE):
                            self.log_file.write(f"    跳过默认docPr名称: '{text_content}'\n")
                        else:
                            position = self._get_textbox_position(elem, img_width, img_height, drawing_elem)
                            positions.append(position)
                            textboxes.append({
                                'elem': elem,
                                'html': escape(text_content.strip())
                            })
                            self.log_file.write(f"    提取docPr文本: '{text_content}', 位置: {position}\n")
                
                # 方法4: 查找a:t（DrawingML文本）
                elif elem.tag == DocxUtils.A_NS + 't':
                    if elem.text and elem.text.strip():
                        text_content = elem.text.strip()
                        parent = self._find_parent_with_tag(elem, [DocxUtils.W_NS + 'txbx', DocxUtils.W_NS + 'txbxContent'])
                        if parent:
                            pass
                        else:
                            position = self._get_textbox_position(elem, img_width, img_height, drawing_elem)
                            positions.append(position)
                            textboxes.append({
                                'elem': elem,
                                'html': escape(text_content)
                            })
                            self.log_file.write(f"    提取DrawingML文本: '{text_content}', 位置: {position}\n")
                
                # 方法5: 查找DrawingML形状 a:sp + a:txBody 文本（Word形状文本框）
                elif elem.tag == DocxUtils.A_NS + 'sp':
                    spPr = elem.find(f'{DocxUtils.A_NS}spPr')
                    shape_css = self._get_shape_style_css(spPr) if spPr is not None else ""
                    txBody = elem.find(f'.//{DocxUtils.A_NS}txBody')
                    if txBody is not None:
                        text_parts = []
                        for t in txBody.findall(f'.//{DocxUtils.A_NS}t'):
                            if t.text:
                                text_parts.append(t.text)
                        inner_text = ''.join(text_parts).strip()
                        inner_html = escape(inner_text)
                        position = self._get_textbox_position(elem, img_width, img_height, drawing_elem)
                        positions.append(position)
                        textboxes.append({
                            'elem': elem,
                            'html': inner_html,
                            'border_css': shape_css
                        })
                        self.log_file.write(f"    提取形状文本框: 文本长度={len(inner_html)}, 位置: {position}\n")
                    else:
                        position = self._get_textbox_position(elem, img_width, img_height, drawing_elem)
                        positions.append(position)
                        textboxes.append({
                            'elem': elem,
                            'html': '',
                            'border_css': shape_css
                        })
                        self.log_file.write(f"    提取无文本形状: 位置: {position}\n")
                
                # 方法6: 组形状 a:grpSp（将整个组作为一个叠加层，保留宽高与边框）
                elif elem.tag == DocxUtils.A_NS + 'grpSp':
                    position = self._get_textbox_position(elem, img_width, img_height, drawing_elem)
                    textboxes.append({
                        'elem': elem,
                        'html': ''
                    })
                    positions.append(position)
                    self.log_file.write(f"    提取组形状grpSp: 位置: {position}\n")
                
                # 方法7: 连接线形状 a:cxnSp（通常无文本，但需保留其边框/线条作为叠加层）
                elif elem.tag == DocxUtils.A_NS + 'cxnSp':
                    position = self._get_textbox_position(elem, img_width, img_height, drawing_elem)
                    textboxes.append({
                        'elem': elem,
                        'html': ''
                    })
                    positions.append(position)
                    self.log_file.write(f"    提取连接形状cxnSp: 位置: {position}\n")
                
                # 方法8: WordprocessingShape 形状 wps:wsp（支持有文本和无文本形状）
                elif elem.tag == DocxUtils.WPS_NS + 'wsp':
                    # 提取形状样式
                    spPr = elem.find(f'{DocxUtils.WPS_NS}spPr')
                    style_css = self._get_shape_style_css(spPr) if spPr is not None else ""
                    
                    txBody = elem.find(f'.//{DocxUtils.A_NS}txBody')
                    if txBody is not None:
                        text_parts = []
                        for t in txBody.findall(f'.//{DocxUtils.A_NS}t'):
                            if t.text:
                                text_parts.append(t.text)
                        inner_text = ''.join(text_parts).strip()
                        inner_html = escape(inner_text)
                        # 如果有文本，且有样式，尝试把样式加到inner_html的外层div（如果之后是div的话）
                        # 这里我们把样式存入 textbox dict，在生成叠加层时使用
                    else:
                        inner_html = ''
                    
                    position = self._get_textbox_position(elem, img_width, img_height, drawing_elem)
                    positions.append(position)
                    textboxes.append({
                        'elem': elem,
                        'html': inner_html,
                        'border_css': style_css  # 将样式存入 border_css 字段，会被叠加层渲染逻辑使用
                    })
                    self.log_file.write(f"    提取wps形状: 文本长度={len(inner_html)}, 位置: {position}, 样式: {style_css}\n")
        
        except Exception as e:
            self.log_file.write(f"  提取文本框失败: {e}\n")
            import traceback
            self.log_file.write(f"  异常详情: {traceback.format_exc()}\n")
        
        self.log_file.write(f"    最终提取 {len(textboxes)} 个文本框\n")
        return textboxes, positions
    
    def _get_shape_style_css(self, spPr):
        """
        从spPr元素中提取CSS样式
        """
        styles = []
        try:
            # 1. 填充颜色
            # 只读取 spPr 直接子元素的 solidFill，避免误把 ln 内的线条颜色当作填充色
            solidFill = spPr.find(f'{DocxUtils.A_NS}solidFill')
            if solidFill is not None:
                srgbClr = solidFill.find(f'{DocxUtils.A_NS}srgbClr')
                if srgbClr is not None:
                    val = srgbClr.get('val')
                    if val:
                        styles.append(f'background-color: #{val}')
            
            # 2. 边框
            ln = spPr.find(f'{DocxUtils.A_NS}ln')
            if ln is not None:
                # 宽度
                w = ln.get('w')
                width_px = 1
                if w:
                    width_px = int(DocxUtils.emu_to_pixels(int(w), apply_scale=True))
                    if width_px < 1: width_px = 1
                
                solidFill_ln = ln.find(f'{DocxUtils.A_NS}solidFill')
                noFill_ln = ln.find(f'{DocxUtils.A_NS}noFill')
                
                if noFill_ln is not None:
                    styles.append('border: none')
                else:
                    border_color = None
                    if solidFill_ln is not None:
                        srgbClr_ln = solidFill_ln.find(f'{DocxUtils.A_NS}srgbClr')
                        if srgbClr_ln is not None:
                            val = srgbClr_ln.get('val')
                            if val:
                                border_color = f'#{val}'
                    
                    # 线型
                    border_style = 'solid'
                    prstDash = ln.find(f'{DocxUtils.A_NS}prstDash')
                    if prstDash is not None:
                        val = prstDash.get('val')
                        if val == 'dot': border_style = 'dotted'
                        elif val == 'dash': border_style = 'dashed'
                    
                    if border_color:
                        styles.append(f'border: {width_px}px {border_style} {border_color}')
                    else:
                        styles.append(f'border: {width_px}px {border_style}')

            # 3. 几何形状 (圆角)
            prstGeom = spPr.find(f'{DocxUtils.A_NS}prstGeom')
            if prstGeom is not None:
                prst = prstGeom.get('prst')
                if prst == 'roundRect':
                    styles.append('border-radius: 15px') # 默认圆角
                elif prst == 'ellipse':
                    styles.append('border-radius: 50%')

        except Exception as e:
            self.log_file.write(f"      提取形状样式失败: {e}\n")
            
        return '; '.join(styles)

    def _find_parent_with_tag(self, elem, tag_list):
        parent = elem
        try:
            getparent = parent.getparent
        except Exception:
            getparent = None
        while parent is not None:
            if hasattr(parent, 'tag') and parent.tag in tag_list:
                return parent
            if getparent:
                parent = parent.getparent()
            else:
                break
        return None
    
    def _get_textbox_position(self, textbox_elem, img_width, img_height, drawing_elem=None):
        """
        获取文本框的位置信息
        
        Args:
            textbox_elem: 文本框元素
            img_width: 图片显示宽度（已应用缩放）
            img_height: 图片显示高度（已应用缩放）
            drawing_elem: 外层drawing元素
            
        Returns:
            dict: {'x', 'y', 'relative_from', 'position_mode'}
        """
        position = {'x': 0, 'y': 0, 'relative_from': None, 'position_mode': 'absolute', 'width': None, 'height': None}
        
        try:
            original_img_width, original_img_height = self._get_original_image_size(drawing_elem)
            scale_x = img_width / original_img_width if original_img_width > 0 else 1
            scale_y = img_height / original_img_height if original_img_height > 0 else 1
            
            # 情况1：textbox_elem 本身是一个独立的 w:drawing（与底图属于不同drawing）
            # 这种情况用于段落中的“独立文本框”挂到同段第一张图片上
            if hasattr(textbox_elem, 'tag') and textbox_elem.tag == DocxUtils.W_NS + 'drawing' and drawing_elem is not None and textbox_elem is not drawing_elem:
                try:
                    base_x, base_y = self._get_anchor_offsets(drawing_elem)
                    tb_x, tb_y = self._get_anchor_offsets(textbox_elem)
                    position['x'] = int(tb_x - base_x)
                    position['y'] = int(tb_y - base_y)
                    self.log_file.write(f"        独立drawing锚点差定位: base=({base_x},{base_y}), tb=({tb_x},{tb_y}), rel=({position['x']},{position['y']})\n")
                except Exception:
                    pass
                
                # 尝试读取该文本框自身的宽高（按页面缩放后的像素）
                try:
                    extent = textbox_elem.find(f'.//{DocxUtils.WP_NS}extent')
                    if extent is not None:
                        cx = extent.get('cx')
                        cy = extent.get('cy')
                        if cx:
                            position['width'] = DocxUtils.emu_to_pixels(int(cx), apply_scale=True)
                        if cy:
                            position['height'] = DocxUtils.emu_to_pixels(int(cy), apply_scale=True)
                except Exception:
                    pass
                
                self.log_file.write(f"      最终文本框位置: x={position['x']}, y={position['y']}\n")
                return position
            
            # 情况2：textbox_elem 是底图drawing内部的元素，使用 a:off / a:ext 相对底图定位
            for elem in textbox_elem.iter():
                if elem.tag == DocxUtils.A_NS + 'off':
                    x = elem.get('x', 0)
                    y = elem.get('y', 0)
                    original_x, original_y = 0, 0
                    if x:
                        original_x = DocxUtils.emu_to_pixels(int(x), apply_scale=True)
                        position['x'] = int(original_x * scale_x)
                    if y:
                        original_y = DocxUtils.emu_to_pixels(int(y), apply_scale=True)
                        position['y'] = int(original_y * scale_y)
                    self.log_file.write(f"        从off获取位置(相对底图): original=({original_x}, {original_y}) -> scaled=({position['x']}, {position['y']})\n")
                if elem.tag == DocxUtils.A_NS + 'ext':
                    cx = elem.get('cx')
                    cy = elem.get('cy')
                    if cx:
                        position['width'] = DocxUtils.emu_to_pixels(int(cx), apply_scale=True) * scale_x
                    if cy:
                        position['height'] = DocxUtils.emu_to_pixels(int(cy), apply_scale=True) * scale_y
            
            # 情况3：仍然没有位置信息时，尝试使用锚点差进行兜底（textbox_elem 为某些外层元素）
            if position['x'] == 0 and position['y'] == 0 and drawing_elem is not None:
                try:
                    base_x, base_y = self._get_anchor_offsets(drawing_elem)
                    tb_x, tb_y = 0, 0
                    if hasattr(textbox_elem, 'tag') and textbox_elem.tag == DocxUtils.W_NS + 'drawing':
                        tb_x, tb_y = self._get_anchor_offsets(textbox_elem)
                    if tb_x or tb_y:
                        position['x'] = int(tb_x - base_x)
                        position['y'] = int(tb_y - base_y)
                        self.log_file.write(f"        兜底锚点差定位: x={position['x']}, y={position['y']}\n")
                except Exception:
                    pass
            
        except Exception as e:
            self.log_file.write(f"  获取文本框位置失败: {e}\n")
            import traceback
            self.log_file.write(f"  异常详情: {traceback.format_exc()}\n")
        
        self.log_file.write(f"      最终文本框位置: x={position['x']}, y={position['y']}\n")
        return position
    
    def _composite_textboxes_to_image(self, img, textboxes, positions, img_width, img_height):
        """
        将文本框内容合成到图片上
        
        Args:
            img: PIL Image对象
            textboxes: 文本框内容列表
            positions: 文本框位置列表
            img_width: 图片宽度
            img_height: 图片高度
            
        Returns:
            PIL Image对象
        """
        try:
            if img.mode != 'RGBA':
                img = img.convert('RGBA')
            
            draw = ImageDraw.Draw(img)
            
            # 智能字体大小计算（基于图片尺寸）
            base_font_size = max(14, min(24, int(img_height / 20)))
            
            # 尝试加载中文字体，提供多种回退选项
            font = None
            font_candidates = [
                ("msyh.ttc", base_font_size),        # 微软雅黑
                ("simhei.ttf", base_font_size),      # 黑体
                ("simsun.ttc", base_font_size),      # 宋体
                ("arial.ttf", base_font_size)        # Arial
            ]
            
            for font_name, size in font_candidates:
                try:
                    font = ImageFont.truetype(font_name, size)
                    self.log_file.write(f"    使用字体: {font_name} {size}px\n")
                    break
                except:
                    continue
            
            if font is None:
                font = ImageFont.load_default()
                self.log_file.write(f"    使用默认字体\n")
            
            # 遍历所有文本框
            for textbox_idx, (textbox, pos) in enumerate(zip(textboxes, positions)):
                # 提取文本
                text = ' '.join(textbox).strip()
                
                if text:
                    self.log_file.write(f"    合并文本框 {textbox_idx}: '{text[:30]}...', 位置: x={pos.get('x')}, y={pos.get('y')}\n")
                    
                    # 获取位置
                    x = int(pos.get('x', 20))
                    y = int(pos.get('y', 20))
                    
                    # 边界检查：确保位置在图片范围内
                    x = max(10, min(x, img_width - 100))
                    y = max(10, min(y, img_height - 50))
                    
                    self.log_file.write(f"      调整后位置: x={x}, y={y}\n")
                    
                    # 计算文本边界
                    try:
                        # 使用textbbox（Pillow 8.0+）
                        bbox = draw.textbbox((0, 0), text, font=font)
                        text_width = bbox[2] - bbox[0]
                        text_height = bbox[3] - bbox[1]
                    except AttributeError:
                        # 回退方案：使用textsize（旧版Pillow）
                        try:
                            text_size = draw.textsize(text, font=font)
                            text_width, text_height = text_size
                        except:
                            # 最终回退：估算
                            text_width = len(text) * base_font_size
                            text_height = base_font_size
                    
                    self.log_file.write(f"      文本尺寸: {text_width}x{text_height}\n")
                    
                    padding = 8
                    
                    # 如果文本太宽，调整位置或缩小字体
                    if text_width > img_width - x - padding * 2:
                        if x > text_width + padding * 2:
                            # 文本放左边
                            x = x - text_width - padding * 2
                        else:
                            # 尝试缩小字体
                            smaller_font_size = max(10, base_font_size - 4)
                            try:
                                smaller_font = ImageFont.truetype("msyh.ttc", smaller_font_size)
                                bbox = draw.textbbox((0, 0), text, font=smaller_font)
                                text_width = bbox[2] - bbox[0]
                                text_height = bbox[3] - bbox[1]
                                font = smaller_font
                            except:
                                pass
                    
                    # 背景矩形位置
                    bg_x1 = max(0, x - padding)
                    bg_y1 = max(0, y - padding)
                    bg_x2 = min(img_width, x + text_width + padding)
                    bg_y2 = min(img_height, y + text_height + padding)
                    
                    # 绘制半透明背景
                    bg_color = (255, 255, 255, 220)  # 半透明白色
                    border_color = (0, 0, 0, 255)    # 不透明黑色
                    
                    # 绘制背景
                    draw.rectangle(
                        [bg_x1, bg_y1, bg_x2, bg_y2],
                        fill=bg_color, outline=border_color, width=2
                    )
                    
                    # 绘制文本阴影
                    shadow_color = (50, 50, 50, 180)
                    draw.text((x + 2, y + 2), text, fill=shadow_color, font=font)
                    
                    # 绘制文本
                    draw.text((x, y), text, fill=(0, 0, 0, 255), font=font)
                    
                    self.log_file.write(f"      背景矩形: {bg_x1},{bg_y1} -> {bg_x2},{bg_y2}\n")
            
            self.log_file.write(f"    文本框合成完成，共处理 {len(textboxes)} 个文本框\n")
            return img
            
        except Exception as e:
            self.log_file.write(f"  文本框合成失败: {e}\n")
            import traceback
            self.log_file.write(f"  异常详情: {traceback.format_exc()}\n")
            return img
    
    def _get_list_prefix_html(self, paragraph):
        """获取列表前缀HTML（序号或项目符号）"""
        list_info, level, list_id, list_color, numFmt, bullet_size = self.parse_list_from_xml(paragraph)
        
        if not list_info:
            return ""
            
        prefix_html = ""
        if list_info == 'numbered':
            # 获取编号
            number = self._get_list_number(list_id, level, numFmt)
            if number:
                number_style = []
                if list_color:
                    number_style.append(f'color: {list_color}')
                else:
                    number_style.append('color: inherit')
                # 字号
                number_size_px = self._get_list_number_size_from_xml(paragraph)
                if number_size_px:
                    number_style.append(f'font-size: {number_size_px}px')
                # 间距依据悬挂缩进/左缩进
                hanging_width = 15
                pPr = paragraph._element.pPr
                if pPr is not None:
                    ind = pPr.find(qn('w:ind'))
                    if ind is not None:
                        hanging = ind.get(qn('w:hanging'))
                        if hanging:
                            hanging_twip = int(hanging)
                            hanging_width = max(15, int(DocxUtils.twip_to_pixels(hanging_twip)))
                        else:
                            left = ind.get(qn('w:left'))
                            if left:
                                left_twip = int(left)
                                left_px = int(DocxUtils.twip_to_pixels(left_twip))
                                hanging_width = max(15, left_px // 3)
                
                number_style_str = '; '.join(number_style)
                prefix_html = f'<span style="{number_style_str}">{number}.</span><span style="display:inline-block; width:{hanging_width}px;"></span>'
        elif list_info == 'bulleted':
            # 项目符号
            bullet_char = list_id if list_id else '•'
            bullet_styles = []
            if list_color:
                bullet_styles.append(f'color: {list_color}')
            else:
                bullet_styles.append('color: inherit')
            # 加粗依据XML
            try:
                rPr0 = paragraph.runs[0]._element.rPr if paragraph.runs else None
                if rPr0 is not None:
                    b0 = rPr0.find(qn('w:b'))
                    if b0 is not None and b0.get(qn('w:val')) != '0':
                        bullet_styles.append('font-weight: bold')
            except Exception:
                pass

            if bullet_size:
                bullet_size_px = int(bullet_size * 1.33)
                bullet_styles.append(f'font-size: {bullet_size_px}px')

            # 从段落属性获取首行缩进，如果没有则使用默认值
            pPr = paragraph._element.pPr
            hanging_width = 15  # 默认15px，增加间隔
            if pPr is not None:
                ind = pPr.find(qn('w:ind'))
                if ind is not None:
                    hanging = ind.get(qn('w:hanging'))
                    if hanging:
                        hanging_twip = int(hanging)
                        hanging_width = max(15, int(DocxUtils.twip_to_pixels(hanging_twip)))
                    else:
                        # 如果没有hanging，检查left缩进
                        left = ind.get(qn('w:left'))
                        if left:
                            left_twip = int(left)
                            left_px = int(DocxUtils.twip_to_pixels(left_twip))
                            hanging_width = max(15, left_px // 3)  # 左缩进的1/3作为项目符号间隔

            bullet_style = '; '.join(bullet_styles)
            prefix_html = f'<span style="{bullet_style}">{bullet_char}</span><span style="display:inline-block; width:{hanging_width}px;"></span>'
            
        return prefix_html

    def process_paragraph(self, paragraph, paragraph_index=None, table_context=None):
        """处理段落 - 基于XML解析"""
        # 只记录简要日志，减少I/O
        if paragraph_index and paragraph_index % 100 == 0:
            self.log_file.write(f"处理段落: {paragraph_index}\n")
        
        # 检查是否包含图片
        images = self.extract_images_from_paragraph(paragraph)
        
        html = ""
        
        if images:
            # 图片段落
            self.log_file.write(f"图片段落: {len(images)} 个图片\n")
            style = self._get_paragraph_format_from_xml(paragraph)
            base_style = f"{style}" if style else ""
            container_style = base_style
            
            # 预先提取段落文本内容，便于决定图片的混排位置
            text_content = self.extract_paragraph_text_with_links(paragraph)
            has_text_content = bool(text_content and text_content.strip())
            
            # 检查是否有浮动图片，如果有则添加清除浮动样式
            has_float = False
            for img in images:
                if img.get('wrap_style') and 'float' in img['wrap_style']:
                    has_float = True
                    break
            
            if has_float and has_text_content:
                prefix_html = self._get_list_prefix_html(paragraph)
                html += f'<p style="{base_style}">{prefix_html}{text_content}</p>'
                has_text_content = False 

            if has_float:
                container_style += "; overflow: hidden;"
            
            html += f'<div style="{container_style}">'
            
            # 处理图片
            for i, img in enumerate(images):
                if img.get('type') == 'textbox':
                    # 处理独立文本框
                    tb_styles = []
                    if img.get('width'):
                        tb_styles.append(f'width: {int(img["width"])}px')
                    # height往往不固定，让内容撑开，或者如有明确高度则设置
                    if img.get('height'):
                         # 某些文本框高度可能只是最小高度
                        tb_styles.append(f'min-height: {int(img["height"])}px')
                    
                    if img.get('wrap_style'):
                        tb_styles.append(img['wrap_style'])
                    
                    if img.get('border_css'):
                        tb_styles.append(img['border_css'])
                    
                    # 补充一些基础样式
                    tb_styles.append('overflow: hidden') # 防止溢出
                    
                    tb_style_str = '; '.join(tb_styles)
                    tb_html = img.get('html', '')
                    
                    self.log_file.write(f"  生成独立文本框: style={tb_style_str}\n")
                    html += f'<div style="{tb_style_str}">{tb_html}</div>'

                elif img.get('type') == 'group':
                    # 处理组合容器
                    group_styles = []
                    if img.get('width'):
                        group_styles.append(f'width: {int(img["width"])}px')
                    if img.get('height'):
                        group_styles.append(f'height: {int(img["height"])}px')
                    
                    if img.get('wrap_style'):
                        group_styles.append(img['wrap_style'])
                    
                    group_styles.append('position: relative')
                    group_style_str = '; '.join(group_styles)
                    
                    html += f'<div style="{group_style_str}">'
                    
                    # 渲染叠加层
                    for ov in img.get('overlays', []):
                        ov_styles = ['position: absolute', 'z-index: 1']
                        ov_styles.append(f'left: {int(ov.get("x", 0))}px')
                        ov_styles.append(f'top: {int(ov.get("y", 0))}px')
                        if ov.get('width'):
                            ov_styles.append(f'width: {int(ov["width"])}px')
                        if ov.get('height'):
                            ov_styles.append(f'height: {int(ov["height"])}px')
                        
                        if ov.get('border_css'):
                            ov_styles.append(ov['border_css'])
                            
                        ov_style_str = '; '.join(ov_styles)
                        
                        if ov.get('type') == 'image' and ov.get('filename'):
                            oimg_path = f'./images/{ov["filename"]}'
                            html += f'<img src="{oimg_path}" style="{ov_style_str}" alt="{ov["filename"]}" />'
                        else:
                            if ov.get('html'):
                                html += f'<div style="{ov_style_str}">{ov["html"]}</div>'
                            else:
                                escaped_text = escape(ov.get('text', ''))
                                html += f'<div style="{ov_style_str}">{escaped_text}</div>'
                    
                    html += '</div>'

                else:
                    # 处理图片 (及旧有的叠加逻辑，虽然现在应该都被group接管了，但保留以防万一)
                    img_style = []
                    if img['width']:
                        img_style.append(f'width: {int(img["width"])}px')
                    if img['height']:
                        img_style.append(f'height: {int(img["height"])}px')
                    if img.get('wrap_style'):
                        img_style.append(img['wrap_style'])
                    
                    # 保持仅基于XML的尺寸，不添加静态样式
                    
                    # 若存在叠加文本框，使用定位容器实现覆盖显示
                    if img.get('overlays'):
                        container_styles = []
                        if img.get('width'):
                            container_styles.append(f'width: {int(img["width"])}px')
                        if img.get('height'):
                            container_styles.append(f'height: {int(img["height"])}px')
                        container_styles.append('position: relative')
                        container_str = '; '.join(container_styles)
                        # 背景图片：绝对定位铺满容器，作为底层背景
                        bg_img_style = []
                        if img.get('width'):
                            bg_img_style.append(f'width: {int(img["width"])}px')
                        if img.get('height'):
                            bg_img_style.append(f'height: {int(img["height"])}px')
                        bg_img_style.append('position: absolute')
                        bg_img_style.append('left: 0')
                        bg_img_style.append('top: 0')
                        bg_img_style.append('z-index: 0')
                        img_style_str = '; '.join(bg_img_style)
                        img_path = f"./images/{img['filename']}"
                        html += f'<div style="{container_str}">'
                        html += f'<img src="{img_path}" style="{img_style_str}" alt="{img["filename"]}" />'
                        # 叠加层
                        for ov in img['overlays']:
                            ov_styles = ['position: absolute', 'z-index: 1']
                            ov_styles.append(f'left: {int(ov.get("x", 0))}px')
                            ov_styles.append(f'top: {int(ov.get("y", 0))}px')
                            if ov.get('width'):
                                ov_styles.append(f'width: {int(ov["width"])}px')
                            if ov.get('height'):
                                ov_styles.append(f'height: {int(ov["height"])}px')
                            if ov.get('border_css'):
                                ov_styles.append(ov['border_css'])
                            ov_style_str = '; '.join(ov_styles)
                            if ov.get('type') == 'image' and ov.get('filename'):
                                oimg_path = f'./images/{ov["filename"]}'
                                html += f'<img src="{oimg_path}" style="{ov_style_str}" alt="{ov["filename"]}" />'
                            else:
                                if ov.get('html'):
                                    html += f'<div style="{ov_style_str}">{ov["html"]}</div>'
                                else:
                                    escaped_text = escape(ov.get('text', ''))
                                    html += f'<div style="{ov_style_str}">{escaped_text}</div>'
                        html += '</div>'
                    else:
                        img_style_str = '; '.join(img_style)
                        img_path = f"./images/{img['filename']}"
                        self.log_file.write(f"  生成图片标签: src={img_path}, style={img_style_str}\n")
                        html += f'<img src="{img_path}" style="{img_style_str}" alt="{img["filename"]}" />'
            
            # 添加段落文本（如果有）
            if has_text_content:
                # 尝试获取列表前缀
                prefix_html = self._get_list_prefix_html(paragraph)
                html += f'<div>{prefix_html}{text_content}</div>'
            
            html += '</div>'
            
        elif not paragraph.text.strip():
            # 空段落
            html = ''
            
        else:
            # 优先检查是否为标题段落（这样即使包含TOC书签但明显是标题的段落也能正确处理）
            if self.is_heading_from_xml(paragraph):
                # 标题段落 - 动态生成ID
                style = self._get_paragraph_format_from_xml(paragraph)
                text = self.extract_paragraph_text_with_links(paragraph)
                
                # 如果标题段落存在自动编号/项目符号，补充前缀
                list_info, level, list_id, list_color, numFmt, bullet_size = self.parse_list_from_xml(paragraph)
                prefix_html = ""
                if list_info == 'numbered':
                    number = self._get_list_number(list_id, level, numFmt)
                    if number:
                        number_style = []
                        if list_color:
                            number_style.append(f'color: {list_color}')
                        else:
                            number_style.append('color: inherit')
                        number_size_px = self._get_list_number_size_from_xml(paragraph)
                        if number_size_px:
                            number_style.append(f'font-size: {number_size_px}px')
                        # 间距依据悬挂缩进/左缩进
                        hanging_width = 15
                        pPr = paragraph._element.pPr
                        if pPr is not None:
                            ind = pPr.find(qn('w:ind'))
                            if ind is not None:
                                hanging = ind.get(qn('w:hanging'))
                                if hanging:
                                    hanging_twip = int(hanging)
                                    hanging_width = max(15, int(DocxUtils.twip_to_pixels(hanging_twip, apply_scale=False)))
                                else:
                                    left = ind.get(qn('w:left'))
                                    if left:
                                        left_twip = int(left)
                                        left_px = int(DocxUtils.twip_to_pixels(left_twip, apply_scale=False))
                                        hanging_width = max(15, left_px // 3)
                        number_style_str = '; '.join(number_style)
                        prefix_html = f'<span style="{number_style_str}">{number}.</span><span style="display:inline-block; width:{hanging_width}px;"></span>'
                elif list_info == 'bulleted':
                    bullet_char = list_id if list_id else '•'
                    bullet_styles = []
                    if list_color:
                        bullet_styles.append(f'color: {list_color}')
                    else:
                        bullet_styles.append('color: inherit')
                    try:
                        rPr0 = paragraph.runs[0]._element.rPr if paragraph.runs else None
                        if rPr0 is not None:
                            b0 = rPr0.find(qn('w:b'))
                            if b0 is not None and b0.get(qn('w:val')) != '0':
                                bullet_styles.append('font-weight: bold')
                    except Exception:
                        pass
                    if bullet_size:
                        bullet_size_px = int(bullet_size * 1.33)
                        bullet_styles.append(f'font-size: {bullet_size_px}px')
                    hanging_width = 15
                    pPr = paragraph._element.pPr
                    if pPr is not None:
                        ind = pPr.find(qn('w:ind'))
                        if ind is not None:
                            hanging = ind.get(qn('w:hanging'))
                            if hanging:
                                hanging_twip = int(hanging)
                                hanging_width = max(15, int(DocxUtils.twip_to_pixels(hanging_twip, apply_scale=False)))
                            else:
                                left = ind.get(qn('w:left'))
                                if left:
                                    left_twip = int(left)
                                    left_px = int(DocxUtils.twip_to_pixels(left_twip, apply_scale=False))
                                    hanging_width = max(15, left_px // 3)
                    bullet_style = '; '.join(bullet_styles)
                    prefix_html = f'<span style="{bullet_style}">{bullet_char}</span><span style="display:inline-block; width:{hanging_width}px;"></span>'
                
                # 获取或生成标题ID
                if paragraph._element in self.paragraph_ids:
                    heading_id = self.paragraph_ids[paragraph._element]
                else:
                    # 如果预扫描未找到，则生成新ID
                    heading_id = self._generate_heading_id(text)
                
                # 注册到标题映射
                heading_text_clean = DocxUtils.strip_html_tags(text).strip()
                self.headings_map[heading_text_clean] = heading_id
                
                html = f'<p id="{heading_id}" style="{style}">{prefix_html}{text}</p>'
                self.log_file.write(f"标题: {heading_text_clean[:30]}... -> #{heading_id}\n")
            
            else:
                # 检查是否为目录段落
                is_toc, toc_title, border_style = self._is_toc_paragraph(paragraph)
                self.log_file.write(f"    目录检测结果: is_toc={is_toc}, toc_title='{toc_title}'\n")
                
                if is_toc:
                    # 目录段落 - 生成跳转链接
                    # 查找匹配的标题ID
                    heading_id = self._find_heading_id(toc_title)
                    # 记录目录标题，用于后续标题识别
                    if toc_title:
                        try:
                            self.toc_titles.add(toc_title.strip())
                        except Exception:
                            pass
                    style = self._get_paragraph_format_from_xml(paragraph)
                    
                    self.log_file.write(f"    目录处理: 找到标题ID={heading_id}, 样式='{style}', 边框='{border_style}'\n")
                    
                    # 合并边框样式到段落样式中
                    if border_style:
                        if style:
                            style += '; ' + border_style
                        else:
                            style = border_style

                    if heading_id and self._paragraph_has_hyperlink(paragraph):
                        html = f'<p style="{style}"><a href="#{heading_id}" style="{DocxUtils.normalize_css("")}">{toc_title}</a></p>'
                        self.log_file.write(f"目录链接: {toc_title} -> #{heading_id}\n")
                    else:
                        html = f'<p style="{style}">{toc_title}</p>'
                        self.log_file.write(f"目录标题未匹配或原文无超链接: {toc_title}\n")
                    
                else:
                    # 若非标题且非目录，尝试使用目录集合辅助识别标题
                    try:
                        text_plain = DocxUtils.strip_html_tags(self.extract_paragraph_text_with_links(paragraph)).strip()
                        matched_toc = self._matches_toc_title(text_plain)
                        if matched_toc:
                            style = self._get_paragraph_format_from_xml(paragraph)
                            text = self.extract_paragraph_text_with_links(paragraph)
                            # 标题前缀（若存在）
                            list_info, level, list_id, list_color, numFmt, bullet_size = self.parse_list_from_xml(paragraph)
                            prefix_html = ""
                            if list_info == 'numbered':
                                number = self._get_list_number(list_id, level, numFmt)
                                if number:
                                    number_style = []
                                    if list_color:
                                        number_style.append(f'color: {list_color}')
                                    else:
                                        number_style.append('color: inherit')
                                    number_size_px = self._get_list_number_size_from_xml(paragraph)
                                    if number_size_px:
                                        number_style.append(f'font-size: {number_size_px}px')
                                    hanging_width = 15
                                    pPr2 = paragraph._element.pPr
                                    if pPr2 is not None:
                                        ind2 = pPr2.find(qn('w:ind'))
                                        if ind2 is not None:
                                            hanging = ind2.get(qn('w:hanging'))
                                            if hanging:
                                                hanging_twip = int(hanging)
                                                hanging_width = max(15, int(DocxUtils.twip_to_pixels(hanging_twip)))
                                            else:
                                                left = ind2.get(qn('w:left'))
                                                if left:
                                                    left_twip = int(left)
                                                    left_px = int(DocxUtils.twip_to_pixels(left_twip))
                                                    hanging_width = max(15, left_px // 3)
                                    number_style_str = '; '.join(number_style)
                                    prefix_html = f'<span style="{number_style_str}">{number}.</span><span style="display:inline-block; width:{hanging_width}px;"></span>'
                            elif list_info == 'bulleted':
                                bullet_char = list_id if list_id else '•'
                                bullet_styles = []
                                if list_color:
                                    bullet_styles.append(f'color: {list_color}')
                                else:
                                    bullet_styles.append('color: inherit')
                                try:
                                    rPr0 = paragraph.runs[0]._element.rPr if paragraph.runs else None
                                    if rPr0 is not None:
                                        b0 = rPr0.find(qn('w:b'))
                                        if b0 is not None and b0.get(qn('w:val')) != '0':
                                            bullet_styles.append('font-weight: bold')
                                except Exception:
                                    pass
                                if bullet_size:
                                    bullet_size_px = int(bullet_size * 1.33)
                                    bullet_styles.append(f'font-size: {bullet_size_px}px')
                                hanging_width = 15
                                pPr3 = paragraph._element.pPr
                                if pPr3 is not None:
                                    ind3 = pPr3.find(qn('w:ind'))
                                    if ind3 is not None:
                                        hanging = ind3.get(qn('w:hanging'))
                                        if hanging:
                                            hanging_twip = int(hanging)
                                            hanging_width = max(15, int(DocxUtils.twip_to_pixels(hanging_twip)))
                                        else:
                                            left = ind3.get(qn('w:left'))
                                            if left:
                                                left_twip = int(left)
                                                left_px = int(DocxUtils.twip_to_pixels(left_twip))
                                                hanging_width = max(15, left_px // 3)
                                bullet_style = '; '.join(bullet_styles)
                                prefix_html = f'<span style="{bullet_style}">{bullet_char}</span><span style="display:inline-block; width:{hanging_width}px;"></span>'
                            
                            # 获取或生成标题ID
                            if paragraph._element in self.paragraph_ids:
                                heading_id = self.paragraph_ids[paragraph._element]
                            else:
                                heading_id = self._generate_heading_id(text)
                                
                            heading_text_clean = DocxUtils.strip_html_tags(text).strip()
                            self.headings_map[heading_text_clean] = heading_id
                            html = f'<p id="{heading_id}" style="{style}">{prefix_html}{text}</p>'
                            self.log_file.write(f"目录辅助识别标题: {heading_text_clean[:30]}... -> #{heading_id}\n")
                            return html
                    except Exception:
                        pass
                    
                    # 检查列表
                    list_info, level, list_id, list_color, numFmt, bullet_size = self.parse_list_from_xml(paragraph)
                    
                    if list_info:
                        # 列表段落
                        style = self._get_paragraph_format_from_xml(paragraph)
                        full_text = self.extract_paragraph_text_with_links(paragraph)
                        
                        if list_info == 'numbered':
                            # 获取编号
                            number = self._get_list_number(list_id, level, numFmt)
                            if number:
                                number_style = []
                                if list_color:
                                    number_style.append(f'color: {list_color}')
                                else:
                                    number_style.append('color: inherit')
                                # 字号
                                number_size_px = self._get_list_number_size_from_xml(paragraph)
                                if number_size_px:
                                    number_style.append(f'font-size: {number_size_px}px')
                                # 间距依据悬挂缩进/左缩进
                                hanging_width = 15
                                pPr = paragraph._element.pPr
                                if pPr is not None:
                                    ind = pPr.find(qn('w:ind'))
                                    if ind is not None:
                                        hanging = ind.get(qn('w:hanging'))
                                        if hanging:
                                            hanging_twip = int(hanging)
                                            hanging_width = max(15, int(DocxUtils.twip_to_pixels(hanging_twip)))
                                        else:
                                            left = ind.get(qn('w:left'))
                                            if left:
                                                left_twip = int(left)
                                                left_px = int(DocxUtils.twip_to_pixels(left_twip))
                                                hanging_width = max(15, left_px // 3)
                                
                                number_style_str = '; '.join(number_style)
                                html = f'<p style="{style}"><span style="{number_style_str}">{number}.</span><span style="display:inline-block; width:{hanging_width}px;"></span><span>{full_text}</span></p>'
                            else:
                                html = f'<p style="{style}">{full_text}</p>'
                        else:
                            # 项目符号
                            bullet_char = list_id if list_id else '•'
                            bullet_styles = []
                            if list_color:
                                bullet_styles.append(f'color: {list_color}')
                            else:
                                bullet_styles.append('color: inherit')
                            # 加粗依据XML
                            try:
                                rPr0 = paragraph.runs[0]._element.rPr if paragraph.runs else None
                                if rPr0 is not None:
                                    b0 = rPr0.find(qn('w:b'))
                                    if b0 is not None and b0.get(qn('w:val')) != '0':
                                        bullet_styles.append('font-weight: bold')
                            except Exception:
                                pass

                            if bullet_size:
                                bullet_size_px = int(bullet_size * 1.33)
                                bullet_styles.append(f'font-size: {bullet_size_px}px')

                            # 从段落属性获取首行缩进，如果没有则使用默认值
                            pPr = paragraph._element.pPr
                            hanging_width = 15  # 默认15px，增加间隔
                            if pPr is not None:
                                ind = pPr.find(qn('w:ind'))
                                if ind is not None:
                                    hanging = ind.get(qn('w:hanging'))
                                    if hanging:
                                        hanging_twip = int(hanging)
                                        hanging_width = max(15, int(DocxUtils.twip_to_pixels(hanging_twip)))
                                    else:
                                        # 如果没有hanging，检查left缩进
                                        left = ind.get(qn('w:left'))
                                        if left:
                                            left_twip = int(left)
                                            left_px = int(DocxUtils.twip_to_pixels(left_twip))
                                            hanging_width = max(15, left_px // 3)  # 左缩进的1/3作为项目符号间隔

                            bullet_style = '; '.join(bullet_styles)
                            html = f'<p style="{style}"><span style="{bullet_style}">{bullet_char}</span><span style="display:inline-block; width:{hanging_width}px;"></span>{full_text}</p>'
                        
                    else:
                        # 普通段落
                        style = self._get_paragraph_format_from_xml(paragraph)
                        text = self.extract_paragraph_text_with_links(paragraph)
                        html = f'<p style="{style}">{text}</p>'
        
        return html
    
    def extract_images_from_paragraph(self, paragraph):
        """从段落中提取图片"""
        images = []
        pending_textboxes = []
        self.log_file.write(f"  开始从段落中提取图片: {paragraph.text[:50]}...\n")
        
        for run_idx, run in enumerate(paragraph.runs):
            drawing_elements = run._element.findall(f'.//{DocxUtils.W_NS}drawing')
            self.log_file.write(f"    Run {run_idx}: 找到 {len(drawing_elements)} 个drawing元素\n")
            
            if drawing_elements:
                for drawing_idx, drawing in enumerate(drawing_elements):
                    # 查找图片
                    blip_elements = drawing.findall(f'.//{DocxUtils.A_NS}blip')
                    self.log_file.write(f"      Drawing {drawing_idx}: 找到 {len(blip_elements)} 个blip元素\n")
                    
                        # 如果没有图片，检查是否为独立文本框或形状
                    if not blip_elements:
                        txbx_content = drawing.find(f'.//{DocxUtils.W_NS}txbxContent')
                        if txbx_content is not None:
                            self.log_file.write(f"      Drawing {drawing_idx}: 找到独立文本框，提取内容\n")
                            inner_html_parts = self._traverse_document_body(txbx_content)
                            inner_html = "".join(inner_html_parts)
                            # 获取文本框的样式
                            txbx_style = self._get_txbx_div_style(txbx_content)
                            if inner_html:
                                pending_textboxes.append({
                                    'elem': drawing,
                                    'html': inner_html,
                                    'style': txbx_style
                                })
                            continue
                        found_shape = None
                        found_tag = None
                        for shape_tag in [
                            DocxUtils.WPS_NS + 'wsp',
                            DocxUtils.A_NS + 'sp',
                            DocxUtils.A_NS + 'grpSp',
                            DocxUtils.A_NS + 'cxnSp',
                        ]:
                            found_shape = drawing.find(f'.//{shape_tag}')
                            if found_shape is not None:
                                found_tag = shape_tag
                                break
                        
                        if found_shape is not None:
                            self.log_file.write(f"      Drawing {drawing_idx}: 找到独立形状 {found_tag}，作为叠加层\n")
                            
                            # 尝试提取样式
                            style_css = ""
                            try:
                                if found_tag == DocxUtils.WPS_NS + 'wsp':
                                    spPr = found_shape.find(f'{DocxUtils.WPS_NS}spPr')
                                    if spPr is not None:
                                        style_css = self._get_shape_style_css(spPr)
                                elif found_tag == DocxUtils.A_NS + 'sp':
                                    spPr = found_shape.find(f'{DocxUtils.A_NS}spPr')
                                    if spPr is not None:
                                        style_css = self._get_shape_style_css(spPr)
                            except Exception as e:
                                self.log_file.write(f"      提取独立形状样式失败: {e}\n")

                            # 过滤不可见的形状 (无HTML内容且样式表明不可见)
                            # 只有当包含可见的边框或背景色时才保留
                            is_visible = False
                            if style_css:
                                # 检查是否有可见边框 (包含 'border:' 且不是 'border: none')
                                has_border = 'border:' in style_css and 'border: none' not in style_css
                                # 检查是否有背景色
                                has_bg = 'background-color:' in style_css
                                
                                if has_border or has_bg:
                                    is_visible = True
                            
                            if not is_visible:
                                self.log_file.write(f"      忽略不可见独立形状 (style='{style_css}')\n")
                                continue

                            pending_textboxes.append({
                                'elem': drawing,
                                'html': '',
                                'border_css': style_css
                            })
                            continue
                    
                    for blip_idx, blip in enumerate(blip_elements):
                        r_embed = blip.get(f'{DocxUtils.R_NS}embed')
                        if r_embed:
                            self.log_file.write(f"        Blip {blip_idx}: 图片关系ID: {r_embed}\n")

                            self.image_counter += 1

                            # 提取图片与文本框位置信息
                            image_info = self.extract_image_with_textbox(
                                r_embed, drawing, self.image_counter
                            )

                            if image_info and image_info.get('filename'):
                                self.log_file.write(f"        成功提取图片: {image_info['filename']}, 尺寸: {image_info['width']}x{image_info['height']}\n")
                                wrap_style = self._get_image_wrap_style(drawing)
                                images.append({
                                    'type': 'image',
                                    'filename': image_info['filename'],
                                    'width': int(image_info['width']) if image_info.get('width') else None,
                                    'height': int(image_info['height']) if image_info.get('height') else None,
                                    'wrap_style': wrap_style,
                                    'overlays': image_info.get('overlays', []),
                                    'drawing_elem': drawing
                                })
                            else:
                                self.log_file.write(f"        提取图片失败: {r_embed}\n")
                        else:
                            self.log_file.write(f"        Blip {blip_idx}: 无embed属性\n")
        
        self.log_file.write(f"  段落图片提取完成，共 {len(images)} 个图片\n")
        
        # 合并所有图片和文本框，进行智能分组
        all_items = []
        # 添加图片
        for img in images:
            img['item_type'] = 'image'
            # 确保有wrap_style
            if 'wrap_style' not in img:
                img['wrap_style'] = self._get_image_wrap_style(img['drawing_elem'])
            all_items.append(img)
            
        # 添加文本框
        for tb in pending_textboxes:
            # 文本框需要构造类似的结构以便统一处理
            # 尝试从xfrm提取宽高
            tb_width = 0
            tb_height = 0
            spPr = None
            if tb['elem'].find(f'.//{DocxUtils.WPS_NS}spPr') is not None:
                spPr = tb['elem'].find(f'.//{DocxUtils.WPS_NS}spPr')
            elif tb['elem'].find(f'.//{DocxUtils.A_NS}spPr') is not None:
                spPr = tb['elem'].find(f'.//{DocxUtils.A_NS}spPr')
                
            if spPr is not None:
                xfrm = spPr.find(f'.//{DocxUtils.A_NS}xfrm')
                if xfrm is not None:
                    ext = xfrm.find(f'.//{DocxUtils.A_NS}ext')
                    if ext is not None:
                        cx = int(ext.get('cx') or 0)
                        cy = int(ext.get('cy') or 0)
                        tb_width = DocxUtils.emu_to_pixels(cx, apply_scale=True)
                        tb_height = DocxUtils.emu_to_pixels(cy, apply_scale=True)
            
            # 获取wrap_style
            wrap_style = self._get_image_wrap_style(tb['elem'])
            
            item = {
                'item_type': 'textbox',
                'html': tb['html'],
                'border_css': tb.get('border_css'),
                'drawing_elem': tb['elem'],
                'width': tb_width,
                'height': tb_height,
                'wrap_style': wrap_style
            }
            all_items.append(item)

        # 为所有元素计算绝对位置
        for it in all_items:
            elem = it.get('drawing_elem')
            x, y = self._get_anchor_offsets(elem)
            it['_abs_x'] = x
            it['_abs_y'] = y
            it['_rect'] = (x, y, it.get('width', 0) or 0, it.get('height', 0) or 0)

        # 按照位置排序：先按y（容差10px），再按x
        # 这样可以保证生成的HTML顺序符合视觉顺序（从上到下，从左到右）
        def sort_key(item):
            # 将y坐标分桶，每10px一行，避免微小抖动导致顺序错乱
            y_bucket = item['_abs_y'] // 10
            return (y_bucket, item['_abs_x'])
            
        all_items.sort(key=sort_key)

        # 简单的重叠检测和分组
        groups = []
        processed = [False] * len(all_items)

        for i in range(len(all_items)):
            if processed[i]:
                continue
            
            # 新建组
            current_group = [all_items[i]]
            processed[i] = True
            
            # 检查后续元素是否与当前组重叠
            # 这是一个简化的贪心策略，可能无法处理复杂的传递性重叠，但对文档排版通常足够
            group_changed = True
            while group_changed:
                group_changed = False
                # 计算当前组的包围盒
                min_x = min(it['_rect'][0] for it in current_group)
                min_y = min(it['_rect'][1] for it in current_group)
                max_x = max(it['_rect'][0] + it['_rect'][2] for it in current_group)
                max_y = max(it['_rect'][1] + it['_rect'][3] for it in current_group)
                
                for j in range(len(all_items)):
                    if not processed[j]:
                        # 检查是否重叠
                        r = all_items[j]['_rect']
                        # 宽松一点的重叠检测（允许一点点间隙或误差，比如5px）
                        margin = 5
                        
                        # 矩形重叠条件：不(A在B左 or A在B右 or A在B上 or A在B下)
                        # r: (x, y, w, h) -> (left, top, width, height)
                        item_left = r[0]
                        item_right = r[0] + r[2]
                        item_top = r[1]
                        item_bottom = r[1] + r[3]
                        
                        is_overlapping = not (
                            item_right + margin < min_x or  # Item在Group左边
                            item_left > max_x + margin or   # Item在Group右边
                            item_bottom + margin < min_y or # Item在Group上边
                            item_top > max_y + margin       # Item在Group下边
                        )
                        
                        if is_overlapping:
                            current_group.append(all_items[j])
                            processed[j] = True
                            group_changed = True
                            # 重新计算包围盒（在下一次循环）
                            break
            
            groups.append(current_group)

        # 构建最终结果列表
        final_results = []
        for group in groups:
            if len(group) == 1:
                # 单个元素，直接返回
                it = group[0]
                if it['item_type'] == 'image':
                    final_results.append(it)
                else:
                    final_results.append({
                        'type': 'textbox',
                        'html': it['html'],
                        'style': '', 
                        'border_css': it.get('border_css', ''),
                        'width': it.get('width'),
                        'height': it.get('height'),
                        'wrap_style': it.get('wrap_style')
                    })
            else:
                # 组合元素
                # 计算组包围盒
                min_x = min(it['_abs_x'] for it in group)
                min_y = min(it['_abs_y'] for it in group)
                max_x = max(it['_abs_x'] + (it.get('width', 0) or 0) for it in group)
                max_y = max(it['_abs_y'] + (it.get('height', 0) or 0) for it in group)
                
                container_w = max_x - min_x
                container_h = max_y - min_y
                
                # 优先使用组内最大的图片的wrap_style，或者第一个元素的
                wrap_style = group[0].get('wrap_style', '')
                max_area = -1
                for it in group:
                    if it['item_type'] == 'image':
                        area = (it.get('width', 0) or 0) * (it.get('height', 0) or 0)
                        if area > max_area:
                            max_area = area
                            wrap_style = it.get('wrap_style', '')
                
                container = {
                    'type': 'group',
                    'width': container_w,
                    'height': container_h,
                    'overlays': [],
                    'wrap_style': wrap_style
                }
                
                for it in group:
                    rel_x = it['_abs_x'] - min_x
                    rel_y = it['_abs_y'] - min_y
                    
                    overlay = {
                        'x': int(rel_x),
                        'y': int(rel_y),
                        'width': it['width'],
                        'height': it['height']
                    }
                    
                    if it['item_type'] == 'image':
                        overlay['type'] = 'image'
                        overlay['filename'] = it['filename']
                    elif it['item_type'] == 'textbox':
                        overlay['type'] = 'text'
                        overlay['html'] = it['html']
                        overlay['border_css'] = it.get('border_css')
                        
                    container['overlays'].append(overlay)
                
                final_results.append(container)
        
        return final_results
    
    def _get_anchor_offsets(self, drawing_elem):
        """获取drawing的锚点绝对位置（页面坐标，像素，已按页面缩放）"""
        x_px, y_px = 0, 0
        try:
            anchor = drawing_elem.find(f'.//{DocxUtils.WP_NS}anchor')
            if anchor is not None:
                positionH = anchor.find(f'.//{DocxUtils.WP_NS}positionH')
                if positionH is not None:
                    posOffset = positionH.find(f'.//{DocxUtils.WP_NS}posOffset')
                    if posOffset is not None and posOffset.text:
                        x_px = DocxUtils.emu_to_pixels(int(posOffset.text), apply_scale=True)
                positionV = anchor.find(f'.//{DocxUtils.WP_NS}positionV')
                if positionV is not None:
                    posOffset = positionV.find(f'.//{DocxUtils.WP_NS}posOffset')
                    if posOffset is not None and posOffset.text:
                        y_px = DocxUtils.emu_to_pixels(int(posOffset.text), apply_scale=True)
        except Exception:
            pass
        return x_px, y_px
    
    def _get_shape_border_css(self, elem):
        """提取shape边框样式为CSS（宽度近似1px，颜色取srgbClr）"""
        try:
            for ln in elem.findall(f'.//{DocxUtils.A_NS}ln'):
                color = None
                solid = ln.find(f'.//{DocxUtils.A_NS}solidFill')
                if solid is not None:
                    srgb = solid.find(f'.//{DocxUtils.A_NS}srgbClr')
                    if srgb is not None and srgb.get('val'):
                        color = f'#{srgb.get("val")}'
                if color:
                    return f'border: 1px solid {color}'
            # 兼容wps:spPr中的线条
            for spPr in elem.findall(f'.//{DocxUtils.A_NS}spPr'):
                ln = spPr.find(f'.//{DocxUtils.A_NS}ln')
                if ln is not None:
                    solid = ln.find(f'.//{DocxUtils.A_NS}solidFill')
                    if solid is not None:
                        srgb = solid.find(f'.//{DocxUtils.A_NS}srgbClr')
                        if srgb is not None and srgb.get('val'):
                            return f'border: 1px solid #{srgb.get("val")}'
        except Exception:
            pass
        return None
    
    def process_table(self, table, table_index, table_context=None):
        """处理表格"""
        self.log_file.write(f"\n{'='*80}\n")
        self.log_file.write(f"表格序号: {table_index}\n")
        self.log_file.write(f"{'='*80}\n")
        
        table_style = []
        table_container_style = []
        
        # 读取表格属性
        tblPr = table._element.find(f'.//{DocxUtils.W_NS}tblPr')
        if tblPr is not None:
            # 读取表格宽度
            tblW = tblPr.find(f'{DocxUtils.W_NS}tblW')
            if tblW is not None:
                w_val = tblW.get(f'{DocxUtils.W_NS}w')
                w_type = tblW.get(f'{DocxUtils.W_NS}type')
                if w_val and w_type:
                    if w_type == 'pct':
                        width_percent = int(w_val) / 50
                        table_style.append(f'width:{width_percent}%')
                    elif w_type == 'dxa':
                        width_twip = int(w_val)
                        width_px = int(width_twip / 20 * 1.33)
                        table_style.append(f'width:{width_px}px')
                else:
                    table_style.append('width:100%')
            
            # 读取表格边框
            tblBorders = tblPr.find(f'.//{DocxUtils.W_NS}tblBorders')
            if tblBorders is not None:
                border_styles = []
                for border_name in ['top', 'left', 'bottom', 'right']:
                    border = tblBorders.find(f'{DocxUtils.W_NS}{border_name}')
                    if border is not None:
                        val = border.get(f'{DocxUtils.W_NS}val')
                        if val and val != 'nil':
                            sz = border.get(f'{DocxUtils.W_NS}sz')
                            color = border.get(f'{DocxUtils.W_NS}color')
                            
                            border_style = f'border-{border_name}:'
                            if sz:
                                border_width = int(int(sz) / 8)
                                border_style += f'{border_width}px'
                            else:
                                border_style += '1px'

                            border_style += ' solid'
                            if color and color != 'auto':
                                border_style += f' #{color}'
                            
                            border_styles.append(border_style)
                
                if border_styles:
                    table_style.append('border-collapse:collapse')
                    table_style.extend(border_styles)
            
            # 读取表格对齐方式
            jc = tblPr.find(f'{DocxUtils.W_NS}jc')
            if jc is not None:
                align_val = jc.get(f'{DocxUtils.W_NS}val')
                if align_val == 'center':
                    table_style.append('margin-left:auto; margin-right:auto')
                elif align_val == 'right':
                    table_style.append('margin-left:auto; margin-right:0')
        
        table_style_str = '; '.join(table_style)
        table_container_style_str = '; '.join(table_container_style)
        
        html = f'<div>'
        html += f'<table style="{table_style_str}">'
        
        skip_cells = set()
        rows = list(table.rows)
        for r_i, row in enumerate(rows):
            html += '<tr>'
            num_cells = len(row.cells)
            c_i = 0
            while c_i < num_cells:
                if (r_i, c_i) in skip_cells:
                    c_i += 1
                    continue
                cell = row.cells[c_i]
                cell_style = []
                tcPr = cell._element.find(f'.//{DocxUtils.W_NS}tcPr')
                colspan = 1
                rowspan = 1
                if tcPr is not None:
                    shd = tcPr.find(f'{DocxUtils.W_NS}shd')
                    if shd is not None:
                        fill = shd.get(f'{DocxUtils.W_NS}fill')
                        if fill and fill != 'auto':
                            cell_style.append(f'background-color:#{fill}')
                    vAlign = tcPr.find(f'{DocxUtils.W_NS}vAlign')
                    if vAlign is not None:
                        align_val = vAlign.get(f'{DocxUtils.W_NS}val')
                        if align_val == 'top':
                            cell_style.append('vertical-align:top')
                        elif align_val == 'center':
                            cell_style.append('vertical-align:middle')
                        elif align_val == 'bottom':
                            cell_style.append('vertical-align:bottom')
                    tcMar = tcPr.find(f'{DocxUtils.W_NS}tcMar')
                    if tcMar is not None:
                        for side in ['top', 'bottom', 'left', 'right']:
                            m = tcMar.find(f'{DocxUtils.W_NS}{side}')
                            if m is not None:
                                w = m.get(f'{DocxUtils.W_NS}w')
                                if w:
                                    try:
                                        px = int(DocxUtils.twip_to_pixels(int(w)))
                                        cell_style.append(f'padding-{side}:{int(px)}px')
                                    except Exception:
                                        pass
                    gridSpan = tcPr.find(f'{DocxUtils.W_NS}gridSpan')
                    if gridSpan is not None:
                        val = gridSpan.get(f'{DocxUtils.W_NS}val')
                        if val:
                            try:
                                colspan = max(1, int(val))
                            except Exception:
                                colspan = 1
                    vMerge = tcPr.find(f'{DocxUtils.W_NS}vMerge')
                    if vMerge is not None:
                        v_val = vMerge.get(f'{DocxUtils.W_NS}val')
                        if v_val and v_val != 'restart':
                            c_i += 1
                            continue
                        else:
                            rr = r_i + 1
                            while rr < len(rows):
                                cell2 = rows[rr].cells[c_i]
                                tcPr2 = cell2._element.find(f'.//{DocxUtils.W_NS}tcPr')
                                if tcPr2 is None:
                                    break
                                vMerge2 = tcPr2.find(f'{DocxUtils.W_NS}vMerge')
                                if vMerge2 is None:
                                    break
                                v_val2 = vMerge2.get(f'{DocxUtils.W_NS}val')
                                if v_val2 == 'restart':
                                    break
                                rowspan += 1
                                skip_cells.add((rr, c_i))
                                rr += 1
                for skip_j in range(1, colspan):
                    skip_cells.add((r_i, c_i + skip_j))
                cell_style_str = '; '.join(cell_style) if cell_style else ''
                span_attrs = ''
                if colspan > 1:
                    span_attrs += f' colspan="{colspan}"'
                if rowspan > 1:
                    span_attrs += f' rowspan="{rowspan}"'
                html += f'<td{span_attrs} style="{cell_style_str}">'
                for para in cell.paragraphs:
                    html += self.process_paragraph(para, f"{table_index}-R{r_i+1}C{c_i+1}")
                html += '</td>'
                c_i += 1
            html += '</tr>'
        
        html += '</table></div>'
        return html
    

    def _scan_toc_recursive(self, element):
        """递归扫描文档中的目录条目"""
        from docx.text.paragraph import Paragraph
        
        for child in element:
            if child.tag == DocxUtils.W_NS + 'p':
                try:
                    paragraph = Paragraph(child, self.doc)
                    is_toc, toc_title, _ = self._is_toc_paragraph(paragraph)
                    if is_toc and toc_title:
                        self.toc_titles.add(toc_title.strip())
                except Exception:
                    pass
            elif child.tag == DocxUtils.W_NS + 'tbl':
                for row in child.findall(f'.//{DocxUtils.W_NS}tr'):
                    for cell in row.findall(f'.//{DocxUtils.W_NS}tc'):
                        self._scan_toc_recursive(cell)
            elif child.tag == DocxUtils.W_NS + 'txbxContent':
                self._scan_toc_recursive(child)
            elif len(child) > 0:
                self._scan_toc_recursive(child)

    def _scan_headings_recursive(self, element):
        """递归扫描文档中的标题，预先生成ID"""
        from docx.text.paragraph import Paragraph
        
        for child in element:
            if child.tag == DocxUtils.W_NS + 'p':
                # 创建段落对象
                try:
                    paragraph = Paragraph(child, self.doc)
                    
                    # 跳过目录段落，避免将其识别为标题
                    is_toc, _, _ = self._is_toc_paragraph(paragraph)
                    if is_toc:
                        continue
                        
                    is_heading = self.is_heading_from_xml(paragraph)
                    
                    if not is_heading:
                        # 尝试匹配目录
                        text_plain = DocxUtils.strip_html_tags(self.extract_paragraph_text_with_links(paragraph)).strip()
                        if self._matches_toc_title(text_plain):
                            is_heading = True
                    
                    if is_heading:
                        text = self.extract_paragraph_text_with_links(paragraph)
                        # 生成ID
                        heading_id = self._generate_heading_id(text)
                        
                        # 注册到标题映射
                        heading_text_clean = DocxUtils.strip_html_tags(text).strip()
                        if heading_text_clean:
                            self.headings_map[heading_text_clean] = heading_id
                        
                        # 存储ID供后续使用
                        self.paragraph_ids[child] = heading_id
                except Exception:
                    pass
                    
            elif child.tag == DocxUtils.W_NS + 'tbl':
                # 递归扫描表格
                for row in child.findall(f'.//{DocxUtils.W_NS}tr'):
                    for cell in row.findall(f'.//{DocxUtils.W_NS}tc'):
                        self._scan_headings_recursive(cell)
                        
            elif child.tag == DocxUtils.W_NS + 'txbxContent':
                # 递归扫描文本框
                self._scan_headings_recursive(child)
                
            elif len(child) > 0:
                # 递归扫描其他复合元素
                self._scan_headings_recursive(child)

    def convert(self):
        """转换DOCX为HTML"""
        # 第一遍扫描：预生成标题ID
        self.log_file.write(f"\n{'='*80}\n")
        self.log_file.write(f"预扫描标题...\n")
        self.log_file.write(f"{'='*80}\n")
        try:
            body = self.doc.part.element.body
            
            # 1. 先扫描目录条目
            self._scan_toc_recursive(body)
            self.log_file.write(f"预扫描完成，发现 {len(self.toc_titles)} 个目录条目\n")
            
            # 2. 再扫描标题（包括匹配目录的标题）
            self._scan_headings_recursive(body)
            self.log_file.write(f"预扫描完成，发现 {len(self.paragraph_ids)} 个标题\n")
        except Exception as e:
            self.log_file.write(f"预扫描标题失败: {e}\n")

        html_parts = []
        
        # HTML头部
        html_header = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{self.doc_title}</title>
    <style>
    </style>
</head>
<body style="width:1200px; margin: 0 auto;">
"""
        html_parts.append(html_header)
        
        # 遍历文档body元素
        body = self.doc.part.element.body
        try:
            total_elements = len(body)
        except Exception:
            total_elements = None
        progress = None
        try:
            if total_elements is not None and total_elements > 0:
                progress = tqdm(total=total_elements, desc="转换进度", dynamic_ncols=True, leave=False)
        except Exception:
            progress = None
        html_parts.extend(self._traverse_document_body(body, progress=progress))
        if progress is not None:
            try:
                progress.close()
            except Exception:
                pass
        
        # HTML尾部
        html_footer = """
</body>
</html>
"""
        html_parts.append(html_footer)
        
        # 保存HTML文件
        print("\n保存HTML文件...", file=sys.stderr)
        html_content = "\n".join(html_parts)
        with open(self.output_html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        # 关闭日志文件
        self.log_file.write(f"\n{'='*80}\n")
        self.log_file.write(f"转换完成!\n")
        self.log_file.write(f"{'='*80}\n")
        self.log_file.write(f"输出目录: {os.path.dirname(self.output_html_path)}\n")
        self.log_file.write(f"HTML文件: {os.path.basename(self.output_html_path)}\n")
        self.log_file.write(f"图片数量: {self.image_counter}\n")
        self.log_file.write(f"标题数量: {len(self.headings_map)}\n")
        self.log_file.write(f"{'='*80}\n")
        self.log_file.close()
        
        print(f"\n转换完成!", file=sys.stderr)
        print(f"HTML文件: {os.path.basename(self.output_html_path)}", file=sys.stderr)
        print(f"图片数量: {self.image_counter}", file=sys.stderr)

    def _get_list_number_size_from_xml(self, paragraph):
        """获取列表编号字号（像素）"""
        if not paragraph.runs:
            return None
        rPr = paragraph.runs[0]._element.rPr
        if rPr is not None:
            sz = rPr.find(qn('w:sz'))
            if sz is not None:
                sz_val = sz.get(qn('w:val'))
                if sz_val:
                    font_size_pt = int(sz_val) / 2
                    return int(font_size_pt * 1.33)
        return None
    
    def _get_txbx_div_style(self, txbx_elem):
        parts = []
        try:
            width = None
            height = None
            left = None
            top = None
            for elem in txbx_elem.iter():
                if elem.tag == DocxUtils.A_NS + 'ext':
                    cx = elem.get('cx')
                    cy = elem.get('cy')
                    if cx:
                        width = DocxUtils.emu_to_pixels(int(cx), apply_scale=True)
                    if cy:
                        height = DocxUtils.emu_to_pixels(int(cy), apply_scale=True)
                    break
            for elem in txbx_elem.iter():
                if elem.tag == DocxUtils.A_NS + 'off':
                    x = elem.get('x')
                    y = elem.get('y')
                    if x:
                        left = DocxUtils.emu_to_pixels(int(x), apply_scale=True)
                    if y:
                        top = DocxUtils.emu_to_pixels(int(y), apply_scale=True)
                    break
            if width:
                parts.append(f'width:{int(width)}px')
            if height:
                parts.append(f'height:{int(height)}px')
            if left is not None or top is not None:
                parts.append('position:relative')
                if left is not None:
                    parts.append(f'left:{int(left)}px')
                if top is not None:
                    parts.append(f'top:{int(top)}px')
        except Exception:
            pass
        return '; '.join(parts)
    
    def _get_image_wrap_style(self, drawing_elem):
        try:
            anchor = drawing_elem.find(f'.//{DocxUtils.WP_NS}anchor')
            if anchor is not None:
                wrap_left = anchor.find(f'.//{DocxUtils.WP_NS}wrapSquare')
                if wrap_left is None:
                    wrap_left = anchor.find(f'.//{DocxUtils.WP_NS}wrapTight')
                if wrap_left is None:
                    wrap_left = anchor.find(f'.//{DocxUtils.WP_NS}wrapNone')
                if wrap_left is None:
                    wrap_left = anchor.find(f'.//{DocxUtils.WP_NS}wrapTopBottom')
                positionH = anchor.find(f'.//{DocxUtils.WP_NS}positionH')
                if positionH is not None:
                    align = positionH.find(f'.//{DocxUtils.WP_NS}align')
                    if align is not None and align.text:
                        if align.text == 'left':
                            return 'float: left'
                        elif align.text == 'right':
                            return 'float: right'
                    # 未提供align时，依据posOffset进行推断
                    posOffset = positionH.find(f'.//{DocxUtils.WP_NS}posOffset')
                    if posOffset is not None and posOffset.text:
                        try:
                            offset_px = DocxUtils.emu_to_pixels(int(posOffset.text), apply_scale=True)
                            # 依据页面宽度的一半作为分界线进行左右推断
                            page_width_px = 1200
                            if offset_px >= page_width_px // 2:
                                return 'float: right'
                            else:
                                return 'float: left'
                        except Exception:
                            pass
                # 如果没有align，根据wrap类型返回可用的默认
                if wrap_left is not None:
                    return None
        except Exception:
            pass
        return None


if __name__ == "__main__":
    import sys
    
    print(f"参数数量: {len(sys.argv)}")
    print(f"参数: {sys.argv}")
    
    if len(sys.argv) != 6:
        print("用法: python convert_docx_to_html.py <docx_file> <html_file> <images_dir> <xml_dir> <log_file>")
        sys.exit(1)
    
    docx_file = sys.argv[1]
    html_file = sys.argv[2]
    images_dir = sys.argv[3]
    xml_dir = sys.argv[4]
    log_file = sys.argv[5]
    
    print(f"输入文件: {docx_file}")
    print(f"输出HTML: {html_file}")
    print(f"图片目录: {images_dir}")
    print(f"XML目录: {xml_dir}")
    print(f"日志文件: {log_file}")
    
    try:
        print(f"开始转换: {docx_file} -> {html_file}")
        converter = DocxToHTMLConverter(docx_file, html_file, images_dir, xml_dir, log_file)
        converter.convert()
        print("转换成功完成!")
    except Exception as e:
        print(f"转换失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)