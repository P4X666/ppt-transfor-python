"""
PPT转换与生成项目主程序

提供PPT转JSON和JSON转PPT的完整功能

使用方法:
    python main.py ppt-to-json --input <pptx文件路径> --output <输出目录>
    python main.py json-to-ppt --input <json文件路径> --output <输出pptx路径>
    python main.py convert --input <pptx文件路径> --output <输出目录>

示例:
    python main.py ppt-to-json --input "ppt/UA模版-封面目录结尾.pptx" --output "output"
    python main.py json-to-ppt --input "output/output.json" --output "output/reconstructed.pptx"
    python main.py convert --input "ppt/UA模版-封面目录结尾.pptx" --output "output"
"""

import os
import sys
import argparse
import logging
from pathlib import Path

from core.ppt_to_json import PPTToJsonConverter, convert_ppt_to_json
from core.json_to_ppt import JsonToPPTConverter, convert_json_to_ppt


def setup_logging(level: str = "INFO"):
    """
    配置日志输出
    
    Args:
        level: 日志级别，可选 DEBUG/INFO/WARNING/ERROR
    """
    log_format = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    logging.basicConfig(
        level=getattr(logging, level.upper(), logging.INFO),
        format=log_format,
        handlers=[
            logging.StreamHandler(sys.stdout)
        ]
    )


def cmd_ppt_to_json(args):
    """PPT转JSON命令"""
    input_path = os.path.abspath(args.input)
    output_dir = os.path.abspath(args.output)
    
    if not os.path.exists(input_path):
        print(f"错误: 输入文件不存在: {input_path}")
        return 1
    
    try:
        json_path = convert_ppt_to_json(
            input_path, 
            output_dir, 
            json_filename=args.json_name
        )
        print(f"转换完成!")
        print(f"JSON文件: {json_path}")
        print(f"图片资源: {os.path.join(output_dir, 'assets', 'images')}")
        return 0
    except Exception as e:
        print(f"转换失败: {e}")
        return 1


def cmd_json_to_ppt(args):
    """JSON转PPT命令"""
    input_path = os.path.abspath(args.input)
    output_path = os.path.abspath(args.output)
    assets_dir = os.path.abspath(args.assets) if args.assets else None
    
    if not os.path.exists(input_path):
        print(f"错误: 输入文件不存在: {input_path}")
        return 1
    
    try:
        ppt_path = convert_json_to_ppt(input_path, output_path, assets_dir)
        print(f"转换完成!")
        print(f"PPT文件: {ppt_path}")
        return 0
    except Exception as e:
        print(f"转换失败: {e}")
        return 1


def cmd_convert(args):
    """完整转换命令: PPT -> JSON -> PPT"""
    input_path = os.path.abspath(args.input)
    output_dir = os.path.abspath(args.output)
    
    if not os.path.exists(input_path):
        print(f"错误: 输入文件不存在: {input_path}")
        return 1
    
    try:
        # 步骤1: PPT转JSON
        print("=" * 50)
        print("步骤 1/2: PPT -> JSON")
        print("=" * 50)
        
        json_path = convert_ppt_to_json(
            input_path, 
            output_dir, 
            json_filename="output.json"
        )
        print(f"JSON文件已生成: {json_path}")
        
        # 步骤2: JSON转PPT
        print("\n" + "=" * 50)
        print("步骤 2/2: JSON -> PPT")
        print("=" * 50)
        
        output_pptx = os.path.join(output_dir, "reconstructed.pptx")
        assets_dir = os.path.join(output_dir, "assets", "images")
        ppt_path = convert_json_to_ppt(json_path, output_pptx, assets_dir)
        print(f"PPT文件已生成: {ppt_path}")
        
        print("\n" + "=" * 50)
        print("转换完成!")
        print("=" * 50)
        print(f"原始PPT: {input_path}")
        print(f"JSON文件: {json_path}")
        print(f"还原PPT: {ppt_path}")
        print(f"图片资源: {assets_dir}")
        
        return 0
    except Exception as e:
        print(f"转换失败: {e}")
        import traceback
        traceback.print_exc()
        return 1


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description="PPT转换与生成工具 - 支持PPT与JSON之间的双向转换",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  %(prog)s ppt-to-json --input template.pptx --output ./output
  %(prog)s json-to-ppt --input ./output/output.json --output reconstructed.pptx
  %(prog)s convert --input template.pptx --output ./output
        """
    )
    
    parser.add_argument(
        "--log-level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        default="INFO",
        help="设置日志级别 (默认: INFO)"
    )
    
    subparsers = parser.add_subparsers(dest="command", help="可用命令")
    
    # ppt-to-json 命令
    ppt2json_parser = subparsers.add_parser(
        "ppt-to-json",
        help="将PPT文件转换为JSON格式"
    )
    ppt2json_parser.add_argument(
        "--input", "-i",
        required=True,
        help="输入PPT文件路径"
    )
    ppt2json_parser.add_argument(
        "--output", "-o",
        required=True,
        help="输出目录路径"
    )
    ppt2json_parser.add_argument(
        "--json-name",
        default="output.json",
        help="输出JSON文件名 (默认: output.json)"
    )
    
    # json-to-ppt 命令
    json2ppt_parser = subparsers.add_parser(
        "json-to-ppt",
        help="将JSON文件转换为PPT格式"
    )
    json2ppt_parser.add_argument(
        "--input", "-i",
        required=True,
        help="输入JSON文件路径"
    )
    json2ppt_parser.add_argument(
        "--output", "-o",
        required=True,
        help="输出PPT文件路径"
    )
    json2ppt_parser.add_argument(
        "--assets",
        help="图片资源目录 (默认从JSON路径推断)"
    )
    
    # convert 命令 (完整流程)
    convert_parser = subparsers.add_parser(
        "convert",
        help="完整转换: PPT -> JSON -> PPT"
    )
    convert_parser.add_argument(
        "--input", "-i",
        required=True,
        help="输入PPT文件路径"
    )
    convert_parser.add_argument(
        "--output", "-o",
        required=True,
        help="输出目录路径"
    )
    
    args = parser.parse_args()
    
    # 配置日志
    setup_logging(args.log_level)
    
    if args.command == "ppt-to-json":
        return cmd_ppt_to_json(args)
    elif args.command == "json-to-ppt":
        return cmd_json_to_ppt(args)
    elif args.command == "convert":
        return cmd_convert(args)
    else:
        parser.print_help()
        return 1


if __name__ == "__main__":
    sys.exit(main())
