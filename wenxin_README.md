# 百度文心智能体批量上传工具

## 功能概述

本工具用于从 Excel 表格中读取夸克网盘分享链接，自动提取图片并批量上传到百度文心智能体数据集。

## 工作流程

### 数据处理流程

1. **读取 Excel** - 从 Excel 文件读取短剧标题和夸克网盘链接
2. **提取夸克链接** - 从 B 列文本中提取夸克网盘分享链接
3. **获取图片** - 访问夸克网盘，获取文件夹中的第一张图片
4. **上传图片** - 将图片上传到文心智能体的 BOS 存储
5. **上传数据** - 将标题、链接、图片地址上传到文心智能体数据集
6. **标记状态** - 在 D 列标记"已上传"，支持断点续传

### 跳过机制

- **D 列已标记"已上传"** → 跳过该行
- **C 列已有图片链接** → 直接提交数据，跳过图片上传

## Excel 表格格式

| A 列（标题） | B 列（链接） | C 列（图片地址） | D 列（状态） |
|-------------|-------------|----------------|-------------|
| 短剧名称 | 夸克/百度网盘链接 | 上传后自动填充 | 已上传/留空 |

## 配置文件说明

配置文件：`config.yaml`

```yaml
# Excel 文件路径
excel:
  input_file: "kuake.xlsx"        # Excel 文件名（相对于脚本目录）
  title_column: 1                 # 标题列（A=1）
  url_column: 2                  # 链接列（B=2）
  image_column: 3                # 图片链接列（C=3）

# 文心智能体配置
agent:
  cookie_file: "baiduck.yaml"    # Cookie 文件路径
  dataset_id: "702225"           # 数据集 ID
  file_id: "5291609"             # 文件 ID
  api_url: "https://agents.baidu.com/lingjing/dataset/text/edit/table"

# 网络请求配置
network:
  timeout: 30                     # 请求超时（秒）
  max_retries: 3                 # 最大重试次数
  request_delay: 1                # 请求间隔（秒）

# 日志配置
logging:
  level: "INFO"                  # 日志级别
  log_file: "upload_direct.log"  # 日志文件
```

## 使用方法

### 1. 准备 Excel 文件

创建 Excel 文件，包含以下列：
- A 列：短剧标题
- B 列：夸克网盘分享链接（可包含百度网盘链接，程序会自动提取夸克链接）
- C 列：留空，程序会自动填充图片地址
- D 列：留空或标记"已上传"用于跳过

### 2. 配置 Cookie

在 `baiduck.yaml` 文件中保存百度账号的 Cookie 信息。

### 3. 运行脚本

```bash
python wenxin.py
```

### 4. 查看日志

程序运行日志会保存到 `upload_direct.log` 文件中。

## 依赖库

- `requests` - HTTP 请求库
- `openpyxl` - Excel 操作库
- `Pillow` - 图片处理库（用于转换 webp 格式）
- `PyYAML` - YAML 配置文件解析

安装依赖：
```bash
pip install requests openpyxl Pillow PyYAML
```

## 核心类说明

### QuarkAPI
夸克网盘 API 封装类，负责：
- 提取文本中的夸克分享链接
- 获取分享 token
- 获取文件列表和预览图 URL

### AgentImageUploader
文心智能体图片上传器，负责：
- 下载夸克网盘图片
- 转换图片格式（webp → JPEG）
- 上传图片到文心智能体 BOS

### AgentDataUploader
文心智能体数据上传器，负责：
- 上传单行数据（标题、链接、图片地址）到数据集

## 注意事项

1. **Cookie 有效期** - 百度 Cookie 会过期，如上传失败请更新 Cookie
2. **网盘访问限制** - 夸克网盘分享链接可能有过期时间
3. **请求间隔** - 默认 1 秒间隔，避免请求过快被限制
4. **断点续传** - 程序会在每行处理后保存 Excel，已上传的行会标记跳过
5. **日志记录** - 详细日志有助于排查问题和确认上传状态

## 错误处理

程序会捕获并记录以下错误：
- Excel 文件不存在
- Cookie 文件不存在或格式错误
- 夸克链接提取失败
- 图片获取/上传失败
- 文心智能体接口返回错误

## 文件说明

| 文件 | 说明 |
|-----|------|
| `wenxin.py` | 主程序脚本 |
| `config.yaml` | 配置文件 |
| `baiduck.yaml` | Cookie 文件 |
| `kuake.xlsx` | Excel 数据文件 |
| `upload_direct.log` | 运行日志 |
