# 制度编写系统 (Zhidu)

企业制度文档结构化编写 + AI 合规审核 + 一键 Word 导出

基于 **Flask** + **TinyMCE 6** 构建，采用 Apple HIG 设计规范。

---

## 功能概览

### 结构化编写
- 预置 11 个标准章节（目的、适用范围、术语定义、管理要求等）
- 每个章节附带填写提示（悬停 `?` 图标查看）
- 章节标题固定不可编辑，保证文档规范统一

### 富文本编辑
- 基于 TinyMCE 6 的完整编辑器
- 文本格式：加粗、斜体、下划线、删除线
- 字体颜色、背景高亮色
- 段落对齐：左 / 居中 / 右 / 两端对齐
- **首行缩进**：一键 2 字符缩进（自定义按钮）
- 有序列表、无序列表
- 表格：插入/编辑/删除，支持**单元格垂直对齐**（顶部 / 居中 / 底部）
- 图片：上传、插入、支持 PNG / JPG / GIF / BMP / WebP / SVG
- Word 粘贴自动清理（去除 MSO 格式、命名空间标签等杂质）
- 查找替换、全屏编辑、字数统计

### AI 智能审核
- 逐章节独立审核，上下文感知
- 审核维度：内容匹配度、完整性、规范性、可操作性、跨章节一致性、主题对齐
- 审核结果以 Markdown 渲染展示
- 支持两种 API 类型：
  - **OpenAI 兼容**：DeepSeek、OpenAI、Moonshot、通义千问等
  - **Anthropic**：Claude 系列模型

### DOCX 一键导出
- 格式完整保留：字体、字号、对齐、缩进、颜色、高亮、删除线
- 表格列宽自动计算（支持 % 和 px）
- 图片嵌入（默认 5 英寸宽度，居中）
- 支持自定义 Word 模板（封面、页眉页脚等）

### 管理后台
- AI 配置管理：增删改查、激活切换
- 管理员密码修改

---

## 快速开始

### 环境要求

| 依赖 | 版本 | 用途 |
|------|------|------|
| Python | 3.10+ | 后端运行 |
| Node.js | 14+ | 安装 TinyMCE 前端资源 |
| Git | 任意 | 代码管理 |

> 下载地址：[Python](https://www.python.org/downloads/) · [Node.js](https://nodejs.org/) · [Git](https://git-scm.com/)

### 安装步骤

```bash
# 1. 克隆项目
git clone https://github.com/Paulwww47/zhidu.git
cd zhidu

# 2. 创建 Python 虚拟环境并安装依赖
python -m venv venv
# Windows:
venv\Scripts\activate
# Linux/macOS:
# source venv/bin/activate

pip install -r requirements.txt

# 3. 安装前端依赖（TinyMCE）
npm install

# 4. 启动应用
python app.py
```

启动后访问：

| 页面 | 地址 |
|------|------|
| 编辑页面 | http://127.0.0.1:5000 |
| 管理后台 | http://127.0.0.1:5000/admin/login |

> 也可使用 Git Bash 运行 `bash start.sh` 一键启动。

### 首次配置

1. **登录管理后台** — 默认账号 `admin`，密码 `admin123`（请立即修改）
2. **配置 AI 接口** — 在"AI 配置管理"中填入你的 API Key

配置示例：

| 字段 | DeepSeek 示例 | Anthropic 示例 |
|------|--------------|----------------|
| 名称 | DeepSeek V3 | Claude Sonnet |
| API Key | sk-xxxxxxxx | sk-ant-xxxxxxxx |
| 模型 | deepseek-chat | claude-sonnet-4-6-20250514 |
| 接口地址 | https://api.deepseek.com | https://api.anthropic.com |
| 接口类型 | OpenAI 兼容 | Anthropic |

---

## 使用指南

### 基本流程

1. 在顶部导航栏输入制度名称
2. 按章节依次填写内容（可参考 `?` 提示）
3. 每个章节完成后点击**"完成"**按钮，AI 自动审核
4. 根据 AI 建议调整内容（审核结果仅供参考）
5. 全部完成后点击右上角**"导出 DOCX"**

### 文档章节结构

| 序号 | 章节 | 说明 |
|------|------|------|
| 一 | 目的 | 制度制定目的和意义 |
| 二 | 适用范围 | 适用的组织、人员和场景 |
| 三 | 规范性引用文件 | 引用的法规、标准、内部制度 |
| 四 | 术语和定义 | 专业术语解释 |
| 五 | 基本原则 | 核心原则和方针 |
| 六 | 职责 | 各部门/岗位职责分工 |
| 七 | 管理要求 | 具体管理规定和操作要求 |
| 八 | 特殊处理机制 | 异常情况处理方式 |
| 九 | 监管与问责 | 检查监督和违规处罚 |
| 十 | 附则 | 解释权、生效日期等 |
| 十一 | 附录 | 附表、流程图等补充材料 |

### 自定义模板导出

如需导出文档包含封面、页眉页脚等预设格式：

1. 在 Word 中创建模板文档（含封面、页眉、页脚、样式等）
2. 命名为 `template.docx`，放置在项目根目录（与 `app.py` 同级）
3. 导出时系统会自动在模板内容之后追加生成的章节内容

> 如果不放置模板文件，系统将创建空白文档并应用默认样式。

建议模板结构：
```
第 1 页：封面（公司 logo、文档标题、日期等）
第 2 页：目录（可选）
第 3 页：空白页（生成内容从这里开始追加）
```

---

## 项目结构

```
zhidu/
├── app.py                  # Flask 主应用（路由、AI 审核、DOCX 导出）
├── requirements.txt        # Python 依赖
├── package.json            # Node.js 依赖（TinyMCE）
├── start.sh                # 一键启动脚本（Git Bash）
├── .gitignore
├── static/
│   ├── css/
│   │   └── style.css       # 全局样式（Apple HIG 风格）
│   └── js/
│       ├── main.js         # 前端主逻辑（编辑器、AI 审核、导出）
│       └── marked.min.js   # Markdown 渲染库
├── templates/
│   ├── editor.html         # 编辑器主页面
│   ├── admin.html          # 管理后台页面
│   └── admin_login.html    # 管理员登录页面
├── uploads/                # 图片上传目录（运行时自动创建）
└── zhidu.db                # SQLite 数据库（首次启动自动创建）
```

---

## 默认配置

### 编辑器
| 配置项 | 默认值 |
|--------|--------|
| 编辑器高度 | 450px |
| 语言 | 中文 (zh_CN) |
| 首行缩进 | 2em |

### DOCX 导出格式
| 配置项 | 默认值 |
|--------|--------|
| 正文字体 | 宋体 / Times New Roman |
| 正文字号 | 12pt（小四号） |
| 标题字体 | 黑体 12pt 加粗 |
| 段前/段后间距 | 各 1 行 |
| 标题段后间距 | 2 行 |
| 行距 | 单倍 |
| 图片宽度 | 5 英寸，居中 |

### AI 审核
| 配置项 | 默认值 |
|--------|--------|
| Temperature | 0.3 |
| Max Tokens | 1500 |

### 管理后台
| 配置项 | 默认值 |
|--------|--------|
| 管理员用户名 | admin |
| 管理员密码 | admin123 |
| 默认 AI 配置 | DeepSeek V3 (deepseek-chat) |

---

## API 路由

| 路由 | 方法 | 说明 |
|------|------|------|
| `/` | GET | 编辑器主页 |
| `/admin/login` | GET/POST | 管理员登录 |
| `/admin/logout` | GET | 退出登录 |
| `/admin` | GET | 管理面板 |
| `/admin/config/add` | POST | 添加 AI 配置 |
| `/admin/config/activate/<id>` | POST | 激活 AI 配置 |
| `/admin/config/delete/<id>` | POST | 删除 AI 配置 |
| `/admin/password` | POST | 修改管理员密码 |
| `/api/ai-check` | POST | AI 智能审核 |
| `/api/export` | POST | 导出 DOCX |
| `/api/upload-image` | POST | 上传图片 |

---

## 技术栈

| 层级 | 技术 |
|------|------|
| 后端框架 | Flask 3.1 |
| Word 生成 | python-docx 1.2 |
| HTML 解析 | BeautifulSoup4 + lxml |
| AI 接口 | OpenAI SDK + Anthropic SDK |
| 富文本编辑器 | TinyMCE 6 (GPL) |
| UI 框架 | Bootstrap 5 |
| Markdown 渲染 | marked.js |
| 数据库 | SQLite |
| 设计规范 | Apple Human Interface Guidelines |

---

## 常见问题

**Q: 启动报错 `ModuleNotFoundError`**
A: 确认已激活虚拟环境并安装依赖：
```bash
venv\Scripts\activate      # Windows
source venv/bin/activate   # Linux/macOS
pip install -r requirements.txt
```

**Q: TinyMCE 编辑器不显示**
A: 确认已运行 `npm install`，检查 `node_modules/tinymce/` 目录是否存在。

**Q: AI 审核报错**
A: 在管理后台检查 AI 配置，确认 API Key 有效且接口地址正确。

**Q: 导出 DOCX 中文字体异常**
A: 系统使用宋体/黑体，确保系统已安装这些字体（Windows 自带，Linux 需手动安装）。

**Q: Linux 下如何安装中文字体？**
A:
```bash
# Ubuntu/Debian
sudo apt install fonts-wqy-microhei fonts-wqy-zenhei

# CentOS/RHEL
sudo yum install wqy-microhei-fonts wqy-zenhei-fonts
```

---

## 许可证

本项目仅供学习和内部使用。TinyMCE 使用 GPL 许可证。
