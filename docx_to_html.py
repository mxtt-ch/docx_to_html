#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DOCX转HTML入口文件
提供命令行入口和main方法
"""

import os
import sys
from datetime import datetime
from convert_docx_to_html import DocxToHTMLConverter


def validate_input_file(docx_path):
    """
    验证输入文件
    
    Args:
        docx_path: DOCX文件路径
        
    Raises:
        FileNotFoundError: 文件不存在
        ValueError: 文件格式错误或文件过大
    """
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"文件不存在: {docx_path}")
    
    if not docx_path.lower().endswith('.docx'):
        raise ValueError(f"文件格式错误，必须是.docx文件: {docx_path}")
    
    file_size = os.path.getsize(docx_path)
    if file_size > 100 * 1024 * 1024:  # 100MB限制
        raise ValueError(f"文件过大: {file_size / (1024*1024):.1f}MB")


def create_output_dirs(output_dir):
    """
    创建输出目录结构
    
    Args:
        output_dir: 输出目录路径
        
    Returns:
        tuple: (html_file, images_dir, xml_dir, log_file)
    """
    # 创建主输出目录
    os.makedirs(output_dir, exist_ok=True)
    
    # 创建子目录
    images_dir = os.path.join(output_dir, "images")
    xml_dir = os.path.join(output_dir, "xml")
    
    os.makedirs(images_dir, exist_ok=True)
    os.makedirs(xml_dir, exist_ok=True)
    
    # 设置文件路径
    html_file = os.path.join(output_dir, "guide.html")
    log_file = os.path.join(output_dir, "convert_log.txt")
    
    return html_file, images_dir, xml_dir, log_file


def main(docx_file=None, output_dir=None, title=None):
    """
    主函数：转换DOCX文件为HTML
    
    Args:
        docx_file: 输入的DOCX文件路径，如果为None则使用默认值
        output_dir: 输出目录路径，如果为None则自动生成时间戳目录
    """
    # 默认文件名
    if docx_file is None:
        docx_file = "ScholarHub数智服务平台使用指南（2025）.docx"
    
    try:
        # 验证输入文件
        validate_input_file(docx_file)
        
        # 创建输出目录
        if output_dir is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_dir = f"output_{timestamp}"
        
        html_file, images_dir, xml_dir, log_file = create_output_dirs(output_dir)
        
        # 打印开始信息
        print("=" * 80)
        print("DOCX转HTML转换器")
        print("=" * 80)
        print(f"输入文件: {docx_file}")
        print(f"输出目录: {output_dir}")
        print(f"HTML文件: guide.html")
        print(f"图片目录: images/")
        print(f"XML目录: xml/")
        print(f"日志文件: convert_log.txt")
        print("=" * 80)
        print()
        
        # 创建转换器并执行转换
        converter = DocxToHTMLConverter(
            docx_path=docx_file,
            output_html_path=html_file,
            images_dir=images_dir,
            xml_dir=xml_dir,
            log_file=log_file,
            title=title
        )
        converter.convert()
        
        print()
        print("=" * 80)
        print("转换完成!")
        print("=" * 80)
        
    except FileNotFoundError as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)
    except ValueError as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"转换过程中发生错误: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    import argparse
    
    # 命令行参数解析
    parser = argparse.ArgumentParser(description='DOCX转HTML转换器')
    parser.add_argument('-i', '--input', type=str, default=None,
                       help='输入的DOCX文件路径')
    parser.add_argument('-o', '--output', type=str, default=None,
                       help='输出目录路径')
    parser.add_argument('-t', '--title', type=str, default=None,
                       help='HTML标题（默认：ScholarHub数智服务平台使用指南（2025））')
    
    args = parser.parse_args()
    
    # 执行转换
    main(args.input, args.output, args.title)
