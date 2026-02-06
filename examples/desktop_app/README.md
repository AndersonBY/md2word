# md2word 轻量桌面 Demo

这是一个放在 `examples/desktop_app` 的最小可运行示例：
- 后端只用 Python 标准库 + 本项目自身
- 前端用 Tailwind CDN，不需要构建工具
- 提供可视化完整配置界面（文档/图片/表格/样式）

## 运行

```bash
python examples/desktop_app/app.py
```

启动后打开：`http://127.0.0.1:7860`

## 说明

- 支持直接粘贴 Markdown 或加载本地 `.md` 文件
- 可选生成目录 (TOC)
- 点击“生成 Word”会下载 `.docx`
- 配置支持导入/导出 JSON
