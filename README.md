# PDF,Office批量转换为Markdown工具

这是一个用于批量将PDF,office文件转换为Markdown格式的工具，基于[MinerU](https://github.com/opendatalab/MinerU)项目的magic_pdf库实现。

> **重要提示**：首次使用前，请务必按照[模型下载说明](#下载模型文件)下载必要的模型文件。

## 功能特点

- 支持批量处理目录中的所有PDF文件
- 支持递归处理子目录中的PDF文件
- 支持处理单个指定的PDF文件
- 支持将Office文件(doc, docx, ppt, pptx)转换为PDF，再转为Markdown
- 自动识别PDF类型，支持OCR和文本模式
- 生成高质量的Markdown文件
- 提取图片并保存到指定目录
- 支持从环境变量文件读取配置
- 使用文件哈希值检查，避免重复处理相同的文件

## 安装依赖

确保已安装所需的依赖：

```bash
pip install -r requirements.txt
```

### 使用Conda环境（推荐）

本项目提供了两种conda环境配置文件：

1. **完整环境**：包含所有依赖及其确切版本，确保与开发环境完全一致

```bash
# 创建完整conda环境
conda env create -f environment-mac.yml
conda env create -f environment-linux.yml

# 激活环境
conda activate simplemarkdown
```

### Office文件转换依赖

要支持Office文件转换为PDF，需要安装LibreOffice：

- **Ubuntu/Debian**:
  ```bash
  sudo apt-get install libreoffice
  ```

- **macOS**:
  ```bash
  brew install libreoffice
  ```

- **Windows**:
  下载并安装 [LibreOffice](https://www.libreoffice.org/download/download/)

## 下载模型文件

本工具基于magic_pdf库，需要下载相关模型文件才能正常工作。请按照以下步骤下载模型：

```bash
# 安装huggingface_hub
pip install huggingface_hub

# 下载模型下载脚本
wget https://github.com/opendatalab/MinerU/raw/master/scripts/download_models_hf.py -O download_models_hf.py

# 执行下载脚本
python download_models_hf.py
```

下载脚本会自动下载所需的模型文件并配置模型目录。模型配置文件将保存在用户目录下的`magic-pdf.json`文件中。

如果您需要更新已下载的模型，只需重新运行上述下载脚本即可。更多详细信息，请参考[MinerU模型下载文档](https://github.com/opendatalab/MinerU/blob/master/docs/how_to_download_models_en.md)。

## 环境变量配置

程序支持通过`.env`文件配置默认的输入输出目录。创建一个`.env`文件，内容如下：

```
INPUT_FOLDER = ./source
OUTPUT_FOLDER = ./md
TMP_PDF_FOLDER = ./tmp_pdf
HASH_DB_PATH = ./md/file_hashes.json
```

如果不存在`.env`文件或未指定这些变量，程序将使用默认值：
- `INPUT_FOLDER`: `source`
- `OUTPUT_FOLDER`: `md`
- `TMP_PDF_FOLDER`: `tmp_pdf`
- `HASH_DB_PATH`: `md/file_hashes.json`

图片目录默认为`OUTPUT_FOLDER/images`。

## 使用方法

### 基本用法

```bash
python pdf_to_md_converter.py
```

默认情况下，程序将：
- 从环境变量`INPUT_FOLDER`指定的目录读取PDF和Office文件
- 将Office文件转换为PDF并保存到`TMP_PDF_FOLDER`目录
- 将生成的Markdown文件保存到环境变量`OUTPUT_FOLDER`指定的目录
- 将提取的图片保存到`OUTPUT_FOLDER/images`目录
- 转换完成后删除临时PDF文件
- 使用哈希值检查跳过未更改的文件

### 自定义目录

```bash
python pdf_to_md_converter.py --source 你的文件目录 --output 输出目录 --images 图片目录 --tmp-pdf 临时PDF目录 --hash-db 哈希值数据库路径
```

### 递归处理子目录

```bash
python pdf_to_md_converter.py --recursive
# 或使用简写形式
python pdf_to_md_converter.py -r
```

### 处理单个文件

```bash
# 处理PDF文件
python pdf_to_md_converter.py --file /path/to/your/file.pdf

# 处理Office文件
python pdf_to_md_converter.py --file /path/to/your/document.docx
python pdf_to_md_converter.py --file /path/to/your/presentation.pptx
```

### 保留临时PDF文件

```bash
python pdf_to_md_converter.py --keep-tmp
```

### 强制处理所有文件

```bash
python pdf_to_md_converter.py --force
# 或使用简写形式
python pdf_to_md_converter.py -f
```

### 参数说明

- `--source`: 源文件目录，默认为环境变量`INPUT_FOLDER`的值
- `--output`: 输出Markdown文件的目录，默认为环境变量`OUTPUT_FOLDER`的值
- `--images`: 输出图片的目录，默认为`OUTPUT_FOLDER/images`
- `--tmp-pdf`: 临时PDF文件目录，默认为环境变量`TMP_PDF_FOLDER`的值
- `--hash-db`: 哈希值数据库文件路径，默认为环境变量`HASH_DB_PATH`的值
- `--recursive`, `-r`: 是否递归处理子目录中的文件
- `--file`: 指定要处理的单个文件路径
- `--keep-tmp`: 保留临时生成的PDF文件
- `--force`, `-f`: 强制处理所有文件，忽略哈希值检查

## 输出文件

对于每个处理的文件（假设名为`example.pdf`或`example.docx`），将生成以下文件：

- `OUTPUT_FOLDER/example.md`: 转换后的Markdown文件
- `OUTPUT_FOLDER/images/`: 包含提取的图片
- `OUTPUT_FOLDER/file_hashes.json`: 文件哈希值数据库，用于跳过未更改的文件
- `TMP_PDF_FOLDER/example.pdf`: 从Office文件转换的临时PDF文件（如果使用了`--keep-tmp`选项）

## 哈希值检查

程序会为每个处理的文件计算SHA-256哈希值，并将其保存在哈希值数据库中。下次运行时，如果文件的哈希值未变化，程序将跳过该文件的处理，从而提高效率。

如果需要强制处理所有文件，可以使用`--force`或`-f`选项禁用哈希值检查。

## 注意事项

- 确保源目录中包含PDF或Office文件
- 程序会自动创建不存在的目录
- 处理大型文件可能需要较长时间
- Office文件转换需要安装LibreOffice
- **首次使用前必须下载模型文件**，否则程序将无法正常工作
- 模型文件较大（约2GB），下载可能需要一些时间，请确保有足够的磁盘空间和稳定的网络连接

## 致谢

本项目基于[MinerU](https://github.com/opendatalab/MinerU)项目开发，使用了其核心组件magic_pdf库进行PDF解析和转换。感谢OpenDataLab团队开发并开源了这一优秀的工具。

如果您在使用过程中遇到任何与magic_pdf库相关的问题，可以参考[MinerU项目文档](https://github.com/opendatalab/MinerU/tree/master)获取更多信息。