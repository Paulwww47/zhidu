# 制度编写系统

基于 Flask + TinyMCE 的企业制度文档编写与 AI 审核系统。

## 功能特性

- **结构化编写**：预置 11 个标准章节（目的、适用范围、术语定义、管理要求等），每个章节附带填写提示
- **富文本编辑**：基于 TinyMCE 6，支持格式化文本、表格、图片插入、首行缩进、垂直对齐等
- **AI 智能审核**：逐章节 AI 审核，检查合规性、完整性、可操作性等维度（支持 OpenAI 兼容接口和 Anthropic 接口）
- **DOCX 导出**：一键导出为 Word 文档，保留格式、对齐、表格列宽、图片等
- **管理后台**：管理 AI 配置（API Key、模型、接口地址）、修改管理员密码

## 环境要求

- **Python** 3.10+
- **Node.js** 14+（仅用于安装 TinyMCE）
- **操作系统**：Windows 10/11（其他系统可能需要微调启动脚本）

## 部署步骤

以下步骤基于全新 Windows 电脑，从零开始部署。

### 1. 安装基础工具

确保已安装以下软件：

- **Python 3.10+**：[下载地址](https://www.python.org/downloads/)
  安装时务必勾选 **"Add Python to PATH"**
- **Node.js 14+**：[下载地址](https://nodejs.org/)
- **Git**：[下载地址](https://git-scm.com/downloads/win)

安装完成后，打开终端验证：

```bash
python --version
node --version
git --version
```

### 2. 克隆项目

```bash
git clone https://github.com/Paulwww47/zhidu.git
cd zhidu
```

### 3. 创建 Python 虚拟环境并安装依赖

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

### 4. 安装 TinyMCE 前端依赖

```bash
npm install
```

这将在 `node_modules/tinymce/` 中安装 TinyMCE 编辑器资源，项目通过 Flask 路由直接提供服务。

### 5. 启动应用

**方式一：直接运行**

```bash
venv\Scripts\activate
python app.py
```

**方式二：使用启动脚本（Git Bash）**

```bash
bash start.sh
```

启动后控制台会显示：

```
 * Running on http://127.0.0.1:5000
```

### 6. 访问系统

| 页面 | 地址 |
|------|------|
| 编辑页面 | http://127.0.0.1:5000 |
| 管理后台 | http://127.0.0.1:5000/admin/login |

### 7. 首次配置

1. **登录管理后台**
   默认管理员账号：`admin`，密码：`admin123`
   *建议首次登录后立即修改密码*

2. **配置 AI 接口**
   在管理后台的"AI 配置管理"中，修改默认配置的 API Key 为你自己的密钥，或新增配置。

   支持的接口类型：
   - **OpenAI 兼容接口**：DeepSeek、OpenAI、Moonshot、通义千问等
   - **Anthropic 接口**：Claude 系列模型

   配置示例：

   | 字段 | DeepSeek 示例 | Anthropic 示例 |
   |------|--------------|----------------|
   | 名称 | DeepSeek V3 | Claude Sonnet |
   | API Key | sk-xxxxxxxx | sk-ant-xxxxxxxx |
   | 模型 | deepseek-chat | claude-sonnet-4-20250514 |
   | 接口地址 | https://api.deepseek.com | https://api.anthropic.com |
   | 接口类型 | OpenAI 兼容 | Anthropic |

## 项目结构

```
zhidu/
├── app.py                  # Flask 主应用（路由、AI 审核、DOCX 导出）
├── requirements.txt        # Python 依赖
├── package.json            # Node.js 依赖（TinyMCE）
├── start.sh                # 启动脚本（Git Bash）
├── static/
│   ├── css/style.css       # Apple HIG 风格样式
│   └── js/
│       ├── main.js         # 前端主逻辑
│       └── marked.min.js   # Markdown 渲染库
├── templates/
│   ├── editor.html         # 编辑页面
│   ├── admin.html          # 管理后台
│   └── admin_login.html    # 管理员登录
├── uploads/                # 图片上传目录（自动创建）
└── zhidu.db                # SQLite 数据库（自动创建）
```

## 使用说明

### 基本流程

1. 在顶部输入制度名称
2. 按章节依次填写内容（可参考每个章节旁的 `?` 提示）
3. 每个章节填写完成后，点击"完成"按钮进行 AI 审核
4. AI 审核结果仅供参考，请根据实际需求调整
5. 全部章节完成后，点击右上角"导出 DOCX"生成 Word 文档

### 使用自定义模板（可选）

如果你希望导出的文档包含封面、页眉、页脚等预设格式，可以使用模板功能：

1. **准备模板文档**
   - 在 Word 中创建一个包含封面、页眉、页脚、样式等的文档
   - 将文档命名为 `template.docx`
   - 放置在项目根目录（与 `app.py` 同级）

2. **导出行为**
   - 如果存在 `template.docx`，系统会打开该模板，并将生成的内容追加到模板后面
   - 如果不存在模板，系统会创建一个空白文档并应用默认样式

3. **注意事项**
   - 模板文档的页眉、页脚、封面等会完整保留
   - 生成的内容会追加在模板现有内容之后
   - 建议模板文档最后留一个空白页，避免内容紧贴封面

**示例模板结构：**
```
第1页：封面（公司 logo、文档标题、日期等）
第2页：目录（可选）
第3页：空白页（生成的内容从这里开始追加）
```

## 常见问题

**Q: 启动报错 `ModuleNotFoundError`**
A: 确认已激活虚拟环境（`venv\Scripts\activate`）并安装了依赖（`pip install -r requirements.txt`）。

**Q: AI 审核报错**
A: 检查管理后台中的 AI 配置，确认 API Key 有效且接口地址正确。

**Q: TinyMCE 编辑器不显示**
A: 确认已运行 `npm install`，检查 `node_modules/tinymce/` 目录是否存在。

**Q: 导出的 DOCX 中文字体异常**
A: 系统使用宋体 / 黑体，确保电脑已安装这些字体（Windows 系统自带）。

## 技术栈

- **后端**：Flask、python-docx、BeautifulSoup、OpenAI SDK、Anthropic SDK
- **前端**：TinyMCE 6（GPL）、Bootstrap 5、marked.js
- **数据库**：SQLite
- **设计规范**：Apple Human Interface Guidelines
