import os
import glob
import argparse
import subprocess
import tempfile
import shutil
import platform
import hashlib
import json
from pathlib import Path
from dotenv import load_dotenv

from magic_pdf.data.data_reader_writer import FileBasedDataWriter, FileBasedDataReader
from magic_pdf.data.dataset import PymuDocDataset
from magic_pdf.model.doc_analyze_by_custom_model import doc_analyze
from magic_pdf.config.enums import SupportedPdfParseMethod

# 支持的Office文件类型
OFFICE_EXTENSIONS = ['.doc', '.docx', '.ppt', '.pptx']

def calculate_file_hash(file_path):
    """
    计算文件的SHA-256哈希值
    
    Args:
        file_path: 文件路径
    
    Returns:
        文件的哈希值字符串
    """
    sha256_hash = hashlib.sha256()
    
    try:
        with open(file_path, "rb") as f:
            # 读取文件块并更新哈希
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
    except Exception as e:
        print(f"计算文件哈希值时出错: {str(e)}")
        return None

def load_hash_database(hash_db_path):
    """
    加载哈希值数据库
    
    Args:
        hash_db_path: 哈希值数据库文件路径
    
    Returns:
        哈希值数据库字典，如果文件不存在则返回空字典
    """
    if os.path.exists(hash_db_path):
        try:
            with open(hash_db_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"加载哈希值数据库时出错: {str(e)}")
            return {}
    return {}

def save_hash_database(hash_db, hash_db_path):
    """
    保存哈希值数据库
    
    Args:
        hash_db: 哈希值数据库字典
        hash_db_path: 哈希值数据库文件路径
    """
    try:
        # 确保目录存在
        os.makedirs(os.path.dirname(hash_db_path), exist_ok=True)
        
        with open(hash_db_path, 'w', encoding='utf-8') as f:
            json.dump(hash_db, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存哈希值数据库时出错: {str(e)}")

def check_libreoffice_installed():
    """
    检查系统是否安装了LibreOffice
    
    Returns:
        tuple: (是否安装, 可执行文件路径)
    """
    # 可能的LibreOffice可执行文件路径
    possible_paths = []
    
    # 根据操作系统添加可能的路径
    system = platform.system()
    if system == "Windows":
        # Windows上的可能路径
        program_files = os.environ.get("ProgramFiles", "C:\\Program Files")
        program_files_x86 = os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)")
        
        possible_paths.extend([
            os.path.join(program_files, "LibreOffice", "program", "soffice.exe"),
            os.path.join(program_files_x86, "LibreOffice", "program", "soffice.exe"),
            os.path.join(program_files, "LibreOffice*", "program", "soffice.exe"),
            os.path.join(program_files_x86, "LibreOffice*", "program", "soffice.exe")
        ])
    elif system == "Darwin":  # macOS
        # macOS上的可能路径
        possible_paths.extend([
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            "/Applications/OpenOffice.app/Contents/MacOS/soffice"
        ])
    else:  # Linux和其他Unix系统
        # Linux上的可能路径
        possible_paths.extend([
            "/usr/bin/soffice",
            "/usr/local/bin/soffice",
            "/opt/libreoffice*/program/soffice"
        ])
    
    # 检查命令是否存在于PATH中
    try:
        # 使用which/where命令查找soffice
        if system == "Windows":
            result = subprocess.run(["where", "soffice"], capture_output=True, text=True)
            if result.returncode == 0:
                return True, result.stdout.strip().split("\n")[0]
        else:
            result = subprocess.run(["which", "soffice"], capture_output=True, text=True)
            if result.returncode == 0:
                return True, result.stdout.strip()
    except:
        pass
    
    # 检查可能的路径
    for path in possible_paths:
        if "*" in path:
            # 处理通配符
            for expanded_path in glob.glob(path):
                if os.path.exists(expanded_path) and os.access(expanded_path, os.X_OK):
                    return True, expanded_path
        elif os.path.exists(path) and os.access(path, os.X_OK):
            return True, path
    
    return False, None

def convert_office_to_pdf(office_file, output_dir="tmp_pdf"):
    """
    将Office文件(doc, docx, ppt, pptx)转换为PDF文件
    
    Args:
        office_file: Office文件路径
        output_dir: 输出PDF文件的目录
    
    Returns:
        转换后的PDF文件路径，如果转换失败则返回None
    """
    # 检查LibreOffice是否安装
    libreoffice_installed, libreoffice_path = check_libreoffice_installed()
    if not libreoffice_installed:
        print("错误: 未找到LibreOffice。请安装LibreOffice以支持Office文件转换。")
        print("安装说明:")
        print("  - Ubuntu/Debian: sudo apt-get install libreoffice")
        print("  - macOS: brew install libreoffice")
        print("  - Windows: 下载并安装 https://www.libreoffice.org/download/download/")
        return None
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    # 获取文件名（不含扩展名）和扩展名
    file_name = os.path.basename(office_file)
    name_without_ext, ext = os.path.splitext(file_name)
    ext = ext.lower()  # 转换为小写进行比较
    
    # 输出PDF文件路径
    pdf_file = os.path.join(output_dir, f"{name_without_ext}.pdf")
    
    print(f"正在将 {file_name} 转换为PDF...")
    
    try:
        # 根据文件类型选择转换方法
        if ext in ['.doc', '.docx', '.ppt', '.pptx']:
            # 使用LibreOffice转换Office文档
            cmd = [libreoffice_path, '--headless', '--convert-to', 'pdf', '--outdir', output_dir, office_file]
            process = subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            
            # 输出详细信息以便调试
            print(f"LibreOffice输出: {process.stdout}")
            if process.stderr:
                print(f"LibreOffice错误: {process.stderr}")
        else:
            print(f"不支持的文件类型: {ext}")
            return None
        
        # 检查PDF文件是否成功生成
        if os.path.exists(pdf_file):
            print(f"成功转换: {office_file} -> {pdf_file}")
            return pdf_file
        else:
            print(f"转换失败: {office_file}")
            return None
    except Exception as e:
        print(f"转换 {office_file} 时出错: {str(e)}")
        return None

def convert_pdf_to_md(pdf_path, output_dir="md", image_dir="md/images", hash_db=None, hash_db_path=None):
    """
    将单个PDF文件转换为Markdown文件
    
    Args:
        pdf_path: PDF文件路径
        output_dir: 输出Markdown文件的目录
        image_dir: 输出图片的目录
        hash_db: 哈希值数据库
        hash_db_path: 哈希值数据库文件路径
    
    Returns:
        bool: 转换是否成功
    """
    # 获取文件名（不含扩展名）
    pdf_file_name = os.path.basename(pdf_path)
    name_without_suff = pdf_file_name.split(".")[0]
    
    # 计算文件哈希值
    current_hash = calculate_file_hash(pdf_path)
    if current_hash is None:
        print(f"无法计算文件哈希值，将继续处理: {pdf_file_name}")
    
    # 检查哈希值是否存在且相同
    if hash_db is not None and current_hash is not None:
        if pdf_path in hash_db and hash_db[pdf_path] == current_hash:
            # 检查MD文件是否存在
            md_file_path = os.path.join(output_dir, f"{name_without_suff}.md")
            if os.path.exists(md_file_path):
                print(f"跳过未更改的文件: {pdf_file_name}")
                return True
    
    # 准备环境
    local_image_dir, local_md_dir = image_dir, output_dir
    image_dir_name = os.path.basename(local_image_dir)
    
    os.makedirs(local_image_dir, exist_ok=True)
    os.makedirs(local_md_dir, exist_ok=True)
    
    image_writer = FileBasedDataWriter(local_image_dir)
    md_writer = FileBasedDataWriter(local_md_dir)
    
    # 读取PDF内容
    reader = FileBasedDataReader("")
    pdf_bytes = reader.read(pdf_path)
    
    # 处理PDF
    # 创建数据集实例
    ds = PymuDocDataset(pdf_bytes)
    
    # 推理
    if ds.classify() == SupportedPdfParseMethod.OCR:
        print(f"使用OCR模式处理: {pdf_file_name}")
        infer_result = ds.apply(doc_analyze, ocr=True)
        pipe_result = infer_result.pipe_ocr_mode(image_writer)
    else:
        print(f"使用文本模式处理: {pdf_file_name}")
        infer_result = ds.apply(doc_analyze, ocr=False)
        pipe_result = infer_result.pipe_txt_mode(image_writer)
    
    # 获取并保存Markdown内容
    md_content = pipe_result.get_markdown(image_dir_name)
    pipe_result.dump_md(md_writer, f"{name_without_suff}.md", image_dir_name)
    
    # 保存哈希值
    if hash_db is not None and current_hash is not None:
        hash_db[pdf_path] = current_hash
        if hash_db_path:
            save_hash_database(hash_db, hash_db_path)
    
    print(f"已完成转换: {pdf_file_name} -> {os.path.join(local_md_dir, f'{name_without_suff}.md')}")
    return True

def batch_convert_files(source_dir="source", output_dir="md", image_dir="md/images", recursive=False, tmp_pdf_dir="tmp_pdf", hash_db_path=None):
    """
    批量转换目录中的所有PDF和Office文件为Markdown文件
    
    Args:
        source_dir: 源文件目录
        output_dir: 输出Markdown文件的目录
        image_dir: 输出图片的目录
        recursive: 是否递归处理子目录
        tmp_pdf_dir: 临时存放转换后PDF文件的目录
        hash_db_path: 哈希值数据库文件路径
    """
    # 确保目录存在
    os.makedirs(source_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(image_dir, exist_ok=True)
    os.makedirs(tmp_pdf_dir, exist_ok=True)
    
    # 加载哈希值数据库
    hash_db = load_hash_database(hash_db_path) if hash_db_path else None
    
    # 获取所有PDF和Office文件
    pdf_files = []
    office_files = []
    
    if recursive:
        # 递归处理
        for root, dirs, files in os.walk(source_dir):
            for file in files:
                file_path = os.path.join(root, file)
                _, ext = os.path.splitext(file)
                ext = ext.lower()  # 转换为小写进行比较
                
                if ext == '.pdf':
                    pdf_files.append(file_path)
                elif ext.lower() in OFFICE_EXTENSIONS:  # 确保使用小写比较
                    office_files.append(file_path)
    else:
        # 处理PDF文件
        for pdf_file in glob.glob(os.path.join(source_dir, "*.pdf")):
            pdf_files.append(pdf_file)
        
        # 处理Office文件 - 使用不区分大小写的模式
        for ext in OFFICE_EXTENSIONS:
            # 同时查找大写和小写扩展名
            pattern_lower = os.path.join(source_dir, f"*{ext}")
            pattern_upper = os.path.join(source_dir, f"*{ext.upper()}")
            
            for file_path in glob.glob(pattern_lower):
                office_files.append(file_path)
            
            for file_path in glob.glob(pattern_upper):
                office_files.append(file_path)
    
    print(f"找到 {len(pdf_files)} 个PDF文件和 {len(office_files)} 个Office文件待处理")
    
    # 检查LibreOffice是否安装
    libreoffice_installed = False
    if office_files:
        libreoffice_installed, _ = check_libreoffice_installed()
        if not libreoffice_installed:
            print("警告: 未找到LibreOffice。将跳过Office文件的处理。")
            print("如需处理Office文件，请安装LibreOffice:")
            print("  - Ubuntu/Debian: sudo apt-get install libreoffice")
            print("  - macOS: brew install libreoffice")
            print("  - Windows: 下载并安装 https://www.libreoffice.org/download/download/")
    
    # 处理Office文件，转换为PDF
    converted_pdf_files = []
    if libreoffice_installed and office_files:
        print(f"开始转换 {len(office_files)} 个Office文件为PDF...")
        for office_file in office_files:
            # 如果是递归模式，保持相对路径结构
            if recursive:
                rel_path = os.path.relpath(os.path.dirname(office_file), source_dir)
                if rel_path != '.':
                    current_tmp_pdf_dir = os.path.join(tmp_pdf_dir, rel_path)
                    os.makedirs(current_tmp_pdf_dir, exist_ok=True)
                else:
                    current_tmp_pdf_dir = tmp_pdf_dir
            else:
                current_tmp_pdf_dir = tmp_pdf_dir
            
            pdf_file = convert_office_to_pdf(office_file, current_tmp_pdf_dir)
            if pdf_file:
                converted_pdf_files.append(pdf_file)
    
    # 合并所有PDF文件列表
    all_pdf_files = pdf_files + converted_pdf_files
    
    if not all_pdf_files:
        print(f"在 {source_dir} 目录中未找到PDF或可转换的Office文件")
        return
    
    print(f"开始处理 {len(all_pdf_files)} 个PDF文件...")
    
    # 处理每个PDF文件
    success_count = 0
    skipped_count = 0
    for pdf_file in all_pdf_files:
        try:
            # 如果是递归模式，保持相对路径结构
            if recursive:
                # 对于原始PDF文件，使用相对于source_dir的路径
                # 对于转换后的PDF文件，使用相对于tmp_pdf_dir的路径
                if pdf_file in pdf_files:
                    rel_path = os.path.relpath(os.path.dirname(pdf_file), source_dir)
                else:
                    rel_path = os.path.relpath(os.path.dirname(pdf_file), tmp_pdf_dir)
                
                if rel_path != '.':
                    current_output_dir = os.path.join(output_dir, rel_path)
                    current_image_dir = os.path.join(image_dir, rel_path)
                    os.makedirs(current_output_dir, exist_ok=True)
                    os.makedirs(current_image_dir, exist_ok=True)
                else:
                    current_output_dir = output_dir
                    current_image_dir = image_dir
            else:
                current_output_dir = output_dir
                current_image_dir = image_dir
            
            result = convert_pdf_to_md(pdf_file, current_output_dir, current_image_dir, hash_db, hash_db_path)
            if result:
                # 检查是否是因为文件未更改而跳过
                if hash_db and pdf_file in hash_db:
                    md_file_path = os.path.join(current_output_dir, f"{os.path.basename(pdf_file).split('.')[0]}.md")
                    if os.path.exists(md_file_path):
                        skipped_count += 1
                    else:
                        success_count += 1
                else:
                    success_count += 1
        except Exception as e:
            print(f"处理 {pdf_file} 时出错: {str(e)}")
    
    # 保存更新后的哈希值数据库
    if hash_db_path and hash_db:
        save_hash_database(hash_db, hash_db_path)
    
    print(f"批量转换完成: 成功 {success_count}/{len(all_pdf_files)}, 跳过 {skipped_count} 个未更改的文件")

if __name__ == "__main__":
    # 加载环境变量
    load_dotenv()
    
    # 从环境变量获取默认目录
    default_source_dir = os.getenv("INPUT_FOLDER", "source").strip()
    default_output_dir = os.getenv("OUTPUT_FOLDER", "md").strip()
    default_image_dir = os.path.join(default_output_dir, "images")
    default_tmp_pdf_dir = os.getenv("TMP_PDF_FOLDER", "tmp_pdf").strip()
    default_hash_db_path = "./file_hashes.json"
    
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='批量将PDF和Office文件转换为Markdown文件')
    parser.add_argument('--source', type=str, default=default_source_dir, help=f'源文件目录 (默认: {default_source_dir})')
    parser.add_argument('--output', type=str, default=default_output_dir, help=f'输出Markdown文件的目录 (默认: {default_output_dir})')
    parser.add_argument('--images', type=str, default=default_image_dir, help=f'输出图片的目录 (默认: {default_image_dir})')
    parser.add_argument('--tmp-pdf', type=str, default=default_tmp_pdf_dir, help=f'临时PDF文件目录 (默认: {default_tmp_pdf_dir})')
    parser.add_argument('--hash-db', type=str, default=default_hash_db_path, help=f'哈希值数据库文件路径 (默认: {default_hash_db_path})')
    parser.add_argument('--recursive', '-r', action='store_true', help='是否递归处理子目录')
    parser.add_argument('--file', type=str, help='指定要处理的单个文件路径')
    parser.add_argument('--keep-tmp', action='store_true', help='保留临时生成的PDF文件')
    parser.add_argument('--force', '-f', action='store_true', help='强制处理所有文件，忽略哈希值检查')
    args = parser.parse_args()
    
    source_dir = args.source
    output_dir = args.output
    image_dir = args.images
    tmp_pdf_dir = args.tmp_pdf
    hash_db_path = None if args.force else args.hash_db
    recursive = args.recursive
    specific_file = args.file
    keep_tmp = args.keep_tmp
    
    print(f"输入目录: {source_dir} (来自环境变量: {'是' if source_dir == default_source_dir else '否'})")
    print(f"输出目录: {output_dir} (来自环境变量: {'是' if output_dir == default_output_dir else '否'})")
    print(f"图片目录: {image_dir}")
    print(f"临时PDF目录: {tmp_pdf_dir}")
    print(f"哈希值数据库: {hash_db_path if hash_db_path else '禁用'}")
    print(f"强制处理所有文件: {'是' if args.force else '否'}")
    
    try:
        # 加载哈希值数据库
        hash_db = load_hash_database(hash_db_path) if hash_db_path else None
        
        # 处理单个文件
        if specific_file:
            if not os.path.exists(specific_file):
                print(f"错误: 文件 {specific_file} 不存在")
                exit(1)
            
            file_name = os.path.basename(specific_file)
            _, ext = os.path.splitext(file_name)
            ext = ext.lower()  # 转换为小写进行比较
            
            if ext == '.pdf':
                # 直接处理PDF文件
                print(f"处理单个PDF文件: {specific_file}")
                try:
                    if convert_pdf_to_md(specific_file, output_dir, image_dir, hash_db, hash_db_path):
                        print(f"成功转换文件: {specific_file}")
                    else:
                        print(f"转换文件失败: {specific_file}")
                except Exception as e:
                    print(f"处理 {specific_file} 时出错: {str(e)}")
            elif ext in OFFICE_EXTENSIONS:
                # 检查LibreOffice是否安装
                libreoffice_installed, _ = check_libreoffice_installed()
                if not libreoffice_installed:
                    print(f"错误: 无法处理 {file_name}。未找到LibreOffice。")
                    print("请安装LibreOffice以支持Office文件转换:")
                    print("  - Ubuntu/Debian: sudo apt-get install libreoffice")
                    print("  - macOS: brew install libreoffice")
                    print("  - Windows: 下载并安装 https://www.libreoffice.org/download/download/")
                    exit(1)
                
                # 先转换为PDF，再处理
                print(f"处理单个Office文件: {specific_file}")
                pdf_file = convert_office_to_pdf(specific_file, tmp_pdf_dir)
                if pdf_file:
                    try:
                        if convert_pdf_to_md(pdf_file, output_dir, image_dir, hash_db, hash_db_path):
                            print(f"成功转换文件: {specific_file} -> {pdf_file} -> Markdown")
                        else:
                            print(f"转换文件失败: {pdf_file}")
                    except Exception as e:
                        print(f"处理 {pdf_file} 时出错: {str(e)}")
                    
                    # 清理临时PDF文件
                    if not keep_tmp and os.path.exists(pdf_file):
                        os.remove(pdf_file)
                        print(f"已删除临时PDF文件: {pdf_file}")
            else:
                print(f"不支持的文件类型: {ext}")
                print(f"支持的文件类型: .pdf, {', '.join(OFFICE_EXTENSIONS)}")
                exit(1)
        else:
            # 批量处理目录
            print(f"递归处理: {'是' if recursive else '否'}")
            print(f"保留临时PDF文件: {'是' if keep_tmp else '否'}")
            
            print("开始批量转换文件为Markdown...")
            batch_convert_files(source_dir, output_dir, image_dir, recursive, tmp_pdf_dir, hash_db_path)
            
            # 清理临时PDF目录
            if not keep_tmp and os.path.exists(tmp_pdf_dir):
                shutil.rmtree(tmp_pdf_dir)
                print(f"已删除临时PDF目录: {tmp_pdf_dir}")
        
        print("转换完成！")
    except KeyboardInterrupt:
        print("\n转换被用户中断")
    except Exception as e:
        print(f"转换过程中出错: {str(e)}") 