# Word 敏感词替换工具

一款完全本地运行的 Word 文档脱敏工具，支持智能识别和手动编辑敏感信息。

## 特性

- **智能识别** - 自动识别手机号、身份证、公司名等 15+ 种敏感信息
- **手动编辑** - 支持添加、删除、修改脱敏项
- **实时预览** - 高亮显示脱敏内容，实时查看效果
- **详细统计** - 显示各类敏感信息的数量
- **完全离线** - 所有处理在本地完成，数据不上传
- **现代界面** - 简洁易用

## 快速开始

### 安装依赖

```bash
# 创建虚拟环境
python3 -m venv .venv

# 激活虚拟环境
source .venv/bin/activate  # Mac/Linux
# 或
.venv\Scripts\activate     # Windows

# 安装依赖
pip install -r requirements.txt
```

### 启动服务

```bash
python main.py
```

浏览器会自动打开 http://localhost:8000

## 支持的脱敏类型

- 手机号
- 身份证号
- 公司名称
- 银行名称
- 网址 URL
- 邮箱地址
- 金额
- 日期
- 地址
- 以及更多...

## 技术栈

- **后端**: FastAPI + Python
- **前端**: HTML + CSS + JavaScript (Tailwind CSS)
- **文档处理**: python-docx

## 依赖

- fastapi >= 0.100.0
- uvicorn >= 0.23.0
- python-docx >= 0.8.11
- pandas >= 2.0.0
- openpyxl >= 3.1.0
- python-multipart >= 0.0.6

## 使用方法

1. 启动服务后，在浏览器中打开工具
2. 上传 Word 文档（.docx 格式）
3. 查看自动识别的敏感信息
4. 手动调整脱敏项（可选）
5. 导出脱敏后的文档

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request。
