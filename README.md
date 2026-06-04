# 📝 论文智能写作助手

基于 DeepSeek AI 大模型的学术论文写作辅助工具。

## ✨ 功能特色

- 🎓 **7种论文模板**：本科/硕士/博士毕业论文、期刊论文、课程论文、研究报告、文献综述
- 🤖 **DeepSeek AI 驱动**：使用 deepseek-v4-pro 模型，1M 上下文窗口
- 📐 **公式图片渲染**：LaTeX 数学公式自动渲染为高清 PNG 图片
- 🔄 **5轮去AI化**：通过多轮多风格改写降低 AI 检测率和查重率
- 💬 **意见反馈系统**：随时提出修改意见，AI 根据反馈调整内容
- 📊 **数据驱动写作**：基于用户提供的研究数据生成真实内容
- 📚 **参考文献管理**：支持链接导入，自动格式化 GB/T 7714 标准
- 📄 **双格式导出**：DOCX + PDF（需安装 Word）
- 🔒 **本地安全存储**：所有数据存储在 R 盘，API 密钥加密保存
- 🌐 **联网检测**：自动检测网络连接，离线时保护功能

## 🚀 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 启动应用

```bash
streamlit run app.py
```

或者在命令行中：
```bash
cd R:\ThesisWriter
streamlit run app.py
```

浏览器会自动打开 http://localhost:8501

### 3. 获取 DeepSeek API Key

访问 [platform.deepseek.com](https://platform.deepseek.com) 注册并获取 API Key。

## 📖 使用流程

1. **选择论文类型** → 选择适合的模板
2. **填写论文信息** → 标题、作者、导师、学校等
3. **输入 API 密钥** → 首次使用需验证密钥
4. **生成大纲** → AI 自动生成，可手动调整
5. **输入研究数据** → 上传 Excel/CSV 或手动描述
6. **逐节写作** → AI 逐节生成，实时流式显示
7. **公式管理** → 输入 LaTeX 公式并渲染为图片
8. **参考文献** → 输入链接，系统自动格式化
9. **去AI化处理** → 多轮改写降低 AI 痕迹
10. **导出文档** → 生成 .docx 和 .pdf 文件

## 📁 数据存储

所有数据存储在 `R:\ThesisWriter\data\` 目录下：
- `theses.db` - SQLite 数据库
- `api_key.enc` - 加密的 API 密钥
- `projects/` - 各项目文件夹

## ⚠️ 注意事项

- 需要稳定的网络连接才能使用 AI 功能
- API 调用会产生费用（DeepSeek 价格实惠）
- 生成的论文内容仅供参考，请遵守学术规范
- PDF 导出需要安装 Microsoft Word
