#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DOCX转HTML通用工具方法
包含单位转换、CSS生成、XML处理等通用功能
"""

import re
from docx.oxml.ns import qn


class DocxUtils:
    """DOCX转换通用工具类"""
    
    # 全局缩放因子，默认1.0。由转换器根据DOCX内容区宽度设置为适配1200px。
    SCALE = 1.0

    # XML命名空间常量
    W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    R_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    A_NS = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    WP_NS = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
    PIC_NS = '{http://schemas.openxmlformats.org/drawingml/2006/picture}'
    V_NS = '{urn:schemas-microsoft-com:vml}'
    W14_NS = '{http://schemas.microsoft.com/office/word/2010/wordml}'
    W15_NS = '{http://schemas.microsoft.com/office/word/2012/wordml}'
    WPS_NS = '{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}'
    MC_NS = '{http://schemas.openxmlformats.org/markup-compatibility/2006}'
    
    @staticmethod
    def emu_to_pixels(emu_value, apply_scale=False):
        """
        EMU转像素
        1 EMU = 1/914400 英寸，96 DPI下 1英寸 = 96像素
        
        Args:
            emu_value: EMU值
            apply_scale: 是否应用全局缩放因子，默认为False
        """
        base = int(emu_value * 96 / 914400) if emu_value else 0
        if apply_scale:
            return int(base * DocxUtils.SCALE)
        else:
            return base
    
    @staticmethod
    def emu_to_points(emu_value):
        """
        EMU转磅
        1 EMU = 1/914400 英寸，1英寸 = 72磅
        """
        return emu_value * 72 / 914400 if emu_value else 0
    
    @staticmethod
    def twip_to_pixels(twip_value, apply_scale=False):
        """
        Twip转像素
        1 Twip = 1/20 磅，1磅 ≈ 1.33像素
        
        Args:
            twip_value: Twip值
            apply_scale: 是否应用全局缩放因子，默认为False
        """
        base = twip_value / 20 * 1.33 if twip_value else 0
        if apply_scale:
            return base * DocxUtils.SCALE
        else:
            return base
    
    @staticmethod
    def twip_to_points(twip_value):
        """
        Twip转磅
        1 Twip = 1/20 磅
        """
        return twip_value / 20 if twip_value else 0
    
    @staticmethod
    def number_to_letter(num, lower=True):
        """数字转字母（1->A, 2->B）"""
        result = ""
        while num > 0:
            num -= 1
            result = chr(ord('A' if not lower else 'a') + num % 26) + result
            num //= 26
        return result
    
    @staticmethod
    def number_to_roman(num, lower=True):
        """数字转罗马数字"""
        val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
        syb = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"]
        roman_num = ''
        i = 0
        while num > 0:
            for _ in range(num // val[i]):
                roman_num += syb[i]
                num -= val[i]
            i += 1
        return roman_num.lower() if lower else roman_num
    
    @staticmethod
    def normalize_css(css_str):
        """
        标准化CSS格式
        移除多余空格和分号，统一格式
        """
        if not css_str:
            return ""
        
        # 移除多余的空格和分号
        css_parts = [part.strip() for part in css_str.split(';') if part.strip()]
        normalized_parts = []
        
        for part in css_parts:
            # 确保属性和值之间有一个冒号和空格
            if ':' in part:
                prop, value = part.split(':', 1)
                prop = prop.strip()
                value = value.strip()
                # 移除值中的多余空格
                value = ' '.join(value.split())
                if prop and value:
                    normalized_parts.append(f"{prop}: {value}")
        
        return '; '.join(normalized_parts)
    
    @staticmethod
    def merge_css_with_priority(css_parts):
        """
        合并CSS，处理优先级冲突
        后定义的属性覆盖先定义的
        """
        css_dict = {}
        
        for css_str in css_parts:
            if not css_str:
                continue
            
            # 解析每个CSS属性
            for prop_value in css_str.split(';'):
                if ':' in prop_value:
                    prop, value = prop_value.split(':', 1)
                    prop = prop.strip()
                    value = value.strip()
                    # 后定义的属性覆盖先定义的
                    if prop and value:
                        css_dict[prop] = value
        
        # 重新组合
        return '; '.join([f"{k}: {v}" for k, v in css_dict.items()])
    
    @staticmethod
    def get_paragraph_style_css(pPr, log_file=None, apply_scale=False):
        """从段落属性XML生成CSS样式（完整版）"""
        styles = []

        if pPr is None:
            return ""

        # 对齐方式
        jc = pPr.find(qn('w:jc'))
        if jc is not None:
            align_map = {'center': 'center', 'right': 'right', 'both': 'justify', 'left': 'left'}
            align_val = jc.get(qn('w:val'), 'left')
            if align_val in align_map:
                styles.append(f'text-align: {align_map[align_val]}')

        # 间距
        spacing = pPr.find(qn('w:spacing'))
        if spacing is not None:
            # 段前间距
            before = spacing.get(qn('w:before'))
            if before:
                before_px = int(DocxUtils.twip_to_pixels(int(before), apply_scale=apply_scale))
                styles.append(f'margin-top: {before_px}px')

            # 段后间距
            after = spacing.get(qn('w:after'))
            if after:
                after_px = int(DocxUtils.twip_to_pixels(int(after), apply_scale=apply_scale))
                styles.append(f'margin-bottom: {after_px}px')

            # 行间距
            line_rule = spacing.get(qn('w:lineRule'))
            line_val = spacing.get(qn('w:line'))
            if line_val:
                if line_rule == 'auto':
                    line_height = int(line_val) / 240
                elif line_rule == 'atLeast':
                    line_height = max(1.0, int(line_val) / 240)
                elif line_rule == 'exact':
                    line_height = int(line_val) / 240
                elif line_rule is None:
                    line_height = int(line_val) / 240
                styles.append(f'line-height: {line_height:.2f}')

        # 缩进
        ind = pPr.find(qn('w:ind'))
        if ind is not None:
            # 左缩进
            left = ind.get(qn('w:left'))
            if left:
                left_px = int(DocxUtils.twip_to_pixels(int(left), apply_scale=apply_scale))
                styles.append(f'margin-left: {left_px}px')

            # 右缩进
            right = ind.get(qn('w:right'))
            if right:
                right_px = int(DocxUtils.twip_to_pixels(int(right), apply_scale=apply_scale))
                styles.append(f'margin-right: {right_px}px')

            # 首行缩进
            first_line = ind.get(qn('w:firstLine'))
            if first_line:
                first_line_px = int(DocxUtils.twip_to_pixels(int(first_line), apply_scale=apply_scale))
                styles.append(f'text-indent: {first_line_px}px')

            # 悬挂缩进（负值）
            hanging = ind.get(qn('w:hanging'))
            if hanging:
                hanging_px = int(DocxUtils.twip_to_pixels(int(hanging), apply_scale=apply_scale))
                styles.append(f'text-indent: -{hanging_px}px')

        # 首字下沉（w:indFirstLineChars）
        ind_first_line_chars = pPr.find(qn('w:indFirstLineChars'))
        if ind_first_line_chars is not None:
            chars = ind_first_line_chars.get(qn('w:val'))
            if chars:
                styles.append('first-letter: float: left; font-size: 200%')

        # 文本方向（w:textDirection）
        text_direction = pPr.find(qn('w:textDirection'))
        if text_direction is not None:
            direction = text_direction.get(qn('w:val'))
            if direction == 'btLr':  # 从下到上
                styles.append('writing-mode: vertical-rl; text-orientation: upright')
            elif direction == 'tbRl':  # 从上到下
                styles.append('writing-mode: vertical-rl')

        return DocxUtils.normalize_css('; '.join(styles))
    
    @staticmethod
    def get_run_style_css(rPr, apply_scale=False):
        """从run属性XML生成CSS样式（完整版）"""
        styles = []
        
        if rPr is None:
            return ""
        
        # 颜色
        color = rPr.find(qn('w:color'))
        if color is not None:
            color_val = color.get(qn('w:val'))
            if color_val:
                if color_val == 'auto':
                    styles.append('color: currentColor')
                elif len(color_val) == 6:
                    styles.append(f'color: #{color_val}')
                elif color_val.startswith('themeColor'):
                    # 处理主题颜色
                    theme_color = color_val.replace('themeColor:', '')
                    styles.append(f'color: var(--{theme_color})')
        
        # 字体高亮
        highlight = rPr.find(qn('w:highlight'))
        if highlight is not None:
            highlight_val = highlight.get(qn('w:val'))
            if highlight_val and highlight_val != 'none':
                color_map = {
                    'yellow': '#FFFF00', 'cyan': '#00FFFF', 'magenta': '#FF00FF',
                    'blue': '#0000FF', 'red': '#FF0000', 'green': '#00FF00',
                    'darkBlue': '#00008B', 'darkRed': '#8B0000', 'darkCyan': '#008B8B'
                }
                if highlight_val in color_map:
                    styles.append(f'background-color: {color_map[highlight_val]}')
        
        # 粗体
        b = rPr.find(qn('w:b'))
        if b is not None and b.get(qn('w:val')) != '0':
            styles.append('font-weight: bold')
        
        # 斜体
        i = rPr.find(qn('w:i'))
        if i is not None and i.get(qn('w:val')) != '0':
            styles.append('font-style: italic')
        
        # 下划线
        u = rPr.find(qn('w:u'))
        if u is not None and u.get(qn('w:val')) != '0':
            u_val = u.get(qn('w:val'), 'single')
            underline_map = {
                'single': 'underline', 'double': 'underline double',
                'dotted': 'underline dotted', 'dashed': 'underline dashed'
            }
            styles.append(f'text-decoration: {underline_map.get(u_val, "underline")}')
        
        # 删除线
        strike = rPr.find(qn('w:strike'))
        if strike is not None and strike.get(qn('w:val')) != '0':
            styles.append('text-decoration: line-through')
        
        # 上下标
        vertAlign = rPr.find(qn('w:vertAlign'))
        if vertAlign is not None:
            val = vertAlign.get(qn('w:val'))
            if val == 'superscript':
                styles.append('vertical-align: super; font-size: 0.8em')
            elif val == 'subscript':
                styles.append('vertical-align: sub; font-size: 0.8em')
        
        # 字符间距
        spacing = rPr.find(qn('w:spacing'))
        if spacing is not None:
            spacing_val = spacing.get(qn('w:val'))
            if spacing_val:
                spacing_twip = int(spacing_val)
                spacing_px = int(DocxUtils.twip_to_pixels(spacing_twip, apply_scale=apply_scale))
                styles.append(f'letter-spacing: {spacing_px}px')
        
        # 字符缩放
        w = rPr.find(qn('w:w'))
        if w is not None:
            w_val = w.get(qn('w:val'))
            if w_val:
                scaling = int(w_val) / 100
                styles.append(f'transform: scaleX({scaling})')
        
        # 阴影
        shd = rPr.find(qn('w:shd'))
        if shd is not None:
            fill = shd.get(qn('w:fill'))
            if fill and fill != 'auto':
                styles.append(f'text-shadow: 1px 1px 2px #{fill}')
        
        return DocxUtils.normalize_css('; '.join(styles))
    
    @staticmethod
    def strip_html_tags(text):
        """移除HTML标签，保留纯文本"""
        return re.sub(r'<[^>]+>', '', text)
    
    @staticmethod
    def extract_text_from_xml_element(elem, w_ns):
        """从XML元素中提取文本内容"""
        text_parts = []
        for t in elem.findall(f'.//{w_ns}t'):
            if t.text:
                text_parts.append(t.text)
        return ''.join(text_parts)
