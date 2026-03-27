# Document Processing Program - AI文档规范化自动化工具
# Document Processing Program - AI Document Normalization Automation Tool

## 产品介绍
## Product Introduction

Document Processing Program是一款轻量级的文档规范化自动化工具，专为学生、自由撰稿人和职场人士设计。它可以帮助用户一键排版、自动纠错、格式统一，告别手动整理文档的繁琐。

Document Processing Program is a lightweight document normalization automation tool designed for students, freelance writers, and professionals. It helps users with one-click formatting, automatic error correction, and format unification, eliminating the hassle of manual document organization.

## 核心功能
## Core Features

### 基础功能
### Basic Features
1. **文档一键格式化**：支持Word（docx）、TXT、PDF文本提取，自动统一字体、字号、行距、页边距、段落缩进，贴合海外院校、职场标准格式
2. **标点&语法纠错**：自动修正中英文标点错误、大小写混乱、语法小失误，去除多余空格、空行、乱码
3. **批量文档处理**：一次性导入多个文档，批量完成格式规整，节省大量时间
4. **本地离线运行**：基础格式化、纠错功能无需联网、不调用API，不泄露用户文档
5. **导出适配**：规整后的文档直接导出为docx、PDF，方便直接提交、打印、上传

1. **One-click Document Formatting**：Supports Word (docx), TXT, and PDF text extraction, automatically unifies font, font size, line spacing, margins, and paragraph indentation, conforming to overseas academic and professional standards
2. **Punctuation & Grammar Correction**：Automatically corrects Chinese and English punctuation errors, case confusion, minor grammar mistakes, and removes extra spaces, blank lines, and garbled characters
3. **Batch Document Processing**：Import multiple documents at once, complete batch format normalization, saving significant time
4. **Local Offline Operation**：Basic formatting and error correction functions do not require internet connection or API calls, protecting user document privacy
5. **Export Compatibility**：Normalized documents can be directly exported as docx or PDF, convenient for direct submission, printing, and uploading

### AI增强功能
### AI-Enhanced Features
1. **AI语句润色**：针对课业、职场文案，优化语句通顺度，贴合学术/正式文风
2. **标题层级自动生成**：自动识别文档标题、副标题，生成规范目录
3. **引用格式简易规整**：针对essay、小论文，自动修正APA、MLA基础引用格式
4. **自定义格式模板**：用户可保存专属排版格式，下次一键套用

1. **AI Sentence Polishing**：Optimizes sentence fluency for academic and professional documents, fitting academic/formal writing styles
2. **Automatic Heading Hierarchy Generation**：Automatically identifies document headings and subheadings, generates standardized table of contents
3. **Citation Format Normalization**：Automatically corrects basic APA and MLA citation formats for essays and papers
4. **Custom Format Templates**：Users can save exclusive formatting styles for one-click application next time

### AI模型支持
### AI Model Support
- **本地模型**：Ollama (支持本地离线运行)
- **云端模型**：
  - OpenAI (GPT-4o, GPT-4-turbo, GPT-3.5-turbo)
  - Anthropic Claude (Claude 3 Opus, Claude 3 Sonnet)
  - Google Gemini (Gemini 1.5 Pro, Gemini 1.5 Flash)
  - 通义千问 (开发中)
  - 文心一言 (开发中)

- **Local Models**：Ollama (supports local offline operation)
- **Cloud Models**：
  - OpenAI (GPT-4o, GPT-4-turbo, GPT-3.5-turbo)
  - Anthropic Claude (Claude 3 Opus, Claude 3 Sonnet)
  - Google Gemini (Gemini 1.5 Pro, Gemini 1.5 Flash)
  - Tongyi Qianwen (in development)
  - Wenxin Yiyan (in development)

### 多语言支持
### Multi-language Support
- **界面语言**：中文、繁体中文、English
- **处理语言**：自动检测、中文、繁体中文、English
- **系统语言自动识别**：根据系统语言自动选择初始界面语言

- **Interface Languages**：Chinese, Traditional Chinese, English
- **Processing Languages**：Auto Detect, Chinese, Traditional Chinese, English
- **System Language Auto-detection**：Automatically selects initial interface language based on system language

## 技术栈
## Technology Stack

- **开发语言**：Python 3.10+
- **网页框架**：Streamlit
- **文档处理库**：python-docx、PyPDF2
- **AI集成**：RESTful API调用
- **部署平台**：Streamlit Community Cloud

- **Development Language**：Python 3.10+
- **Web Framework**：Streamlit
- **Document Processing Libraries**：python-docx, PyPDF2
- **AI Integration**：RESTful API calls
- **Deployment Platform**：Streamlit Community Cloud

## 快速开始
## Quick Start

### 本地运行
### Local Run

1. 克隆项目到本地
2. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```
3. 运行应用：
   ```bash
   streamlit run app.py
   ```
4. 访问应用：http://localhost:8508

1. Clone the project to your local machine
2. Install dependencies：
   ```bash
   pip install -r requirements.txt
   ```
3. Run the application：
   ```bash
   streamlit run app.py
   ```
4. Access the application：http://localhost:8508

### 在线访问
### Online Access

项目已部署在Streamlit Community Cloud，可直接通过以下链接访问：
[Document Processing Program - 在线文档规范化工具](https://cleandocs.streamlit.app)

The project is deployed on Streamlit Community Cloud and can be accessed directly through the following link：
[Document Processing Program - Online Document Normalization Tool](https://cleandocs.streamlit.app)

## 使用方法
## Usage

1. **文档格式化**：上传Word或PDF文档，系统会自动统一格式并提供下载
2. **标点&语法纠错**：可直接输入文本或上传文档，系统会自动修正标点和语法错误
3. **批量文档处理**：上传多个文档，系统会批量处理并提供下载
4. **AI语句润色**：输入文本或上传文档，选择AI模型，系统会优化语句通顺度
5. **标题层级生成**：上传文档，系统会自动识别标题层级并生成目录
6. **引用格式规整**：输入文本或上传文档，选择引用格式（APA/MLA），系统会自动修正格式
7. **自定义格式模板**：创建并保存专属排版格式，下次可直接应用

1. **Document Formatting**：Upload Word or PDF documents, the system will automatically unify the format and provide download
2. **Punctuation & Grammar Correction**：Directly input text or upload documents, the system will automatically correct punctuation and grammar errors
3. **Batch Document Processing**：Upload multiple documents, the system will process them in batch and provide downloads
4. **AI Sentence Polishing**：Input text or upload documents, select AI model, the system will optimize sentence fluency
5. **Heading Hierarchy Generation**：Upload documents, the system will automatically identify heading levels and generate table of contents
6. **Citation Format Normalization**：Input text or upload documents, select citation format (APA/MLA), the system will automatically correct the format
7. **Custom Format Templates**：Create and save exclusive formatting styles for direct application next time

## 适用人群
## Target Users

- **海外大学生/研究生**：课程作业、论文、报告、essay格式规范要求严苛
- **自由撰稿人/自媒体博主**：日常产出文稿、文案、专栏文章
- **职场新人/小微企业文员**：整理工作周报、会议纪要、简易合同
- **海外网课学习者/留学生**：各类课业文档繁多，格式混乱

- **Overseas College/graduate Students**：Strict format requirements for course assignments, papers, reports, and essays
- **Freelance Writers/Content Creators**：Daily production of articles, copywriting, and column pieces
- **New Professionals/Small Business Clerks**：Organizing weekly reports, meeting minutes, and simple contracts
- **Overseas Online Course Learners/International Students**：Numerous academic documents with disorganized formats

## 产品优势
## Product Advantages

- **轻量化**：纯网页版，即开即用，无需安装
- **极简操作**：一键处理，无需学习，全程无复杂操作
- **隐私保护**：本地离线运行，不泄露用户文档
- **低成本**：基础功能免费使用，AI功能调用低成本API
- **多语言支持**：支持中文、繁体中文和英文，满足国际化需求
- **AI增强**：集成多种AI模型，提供智能文档处理能力
- **适配Gumroad**：符合Gumroad用户偏爱"小而美、一次性付费"的消费习惯

- **Lightweight**：Pure web version, ready to use, no installation required
- **Minimal Operation**：One-click processing, no learning required, no complex operations throughout
- **Privacy Protection**：Local offline operation, no user document leakage
- **Low Cost**：Basic features free to use, AI features use low-cost APIs
- **Multi-language Support**：Supports Chinese, Traditional Chinese, and English, meeting international needs
- **AI Enhancement**：Integrates multiple AI models, providing intelligent document processing capabilities
- **Gumroad Compatible**：Meets Gumroad users' preference for "small and beautiful, one-time payment" consumption habits

## 部署说明
## Deployment Instructions

项目使用Streamlit Community Cloud免费部署，无需购买服务器，一键上线网页。

1. 注册Streamlit账号
2. 关联GitHub代码仓库
3. 一键部署上线，生成专属网页链接
4. 部署后，应用会自动同步GitHub仓库的更新

The project uses Streamlit Community Cloud for free deployment, no server purchase required, one-click online launch.

1. Register a Streamlit account
2. Link GitHub code repository
3. One-click deployment online, generate exclusive web link
4. After deployment, the application will automatically sync GitHub repository updates

## 配置说明
## Configuration Instructions

### Streamlit配置
### Streamlit Configuration

项目包含以下配置文件：
- `.streamlit/config.toml`：配置Streamlit服务器设置，包括端口、地址和CORS设置

The project includes the following configuration files：
- `.streamlit/config.toml`：Configures Streamlit server settings, including port, address, and CORS settings

### 环境变量
### Environment Variables

运行应用时，系统会使用以下默认配置：
- Ollama API地址：http://localhost:11434
- 默认Ollama模型：llama3

When running the application, the system uses the following default configurations：
- Ollama API address：http://localhost:11434
- Default Ollama model：llama3

## 许可证
## License

MIT License