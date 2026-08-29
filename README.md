# PPT 图片生成工具

一个本地网页工具：你逐页描述 PPT 内容 → DeepSeek 自动优化成高质量生图提示词 → 调用你自己的 `gpt-image-2` Key 生成 16:9 页面图片 → 自动拼成 `.pptx` 并保留原图。

## 你需要准备

| 用途 | 说明 |
| --- | --- |
| DeepSeek API Key | 用于优化提示词（模型 `deepseek-chat`） |
| OpenAI API Key | 用于生成图片（模型 `gpt-image-2`，需完成账号图片模型资质验证） |

两个 Key 都只在**你自己的浏览器本地**使用，不会上传到任何第三方。

## 一键启动（Windows）

双击运行 `start.bat`，首次会自动创建虚拟环境并安装依赖，然后打开 http://127.0.0.1:5000 。

## 手动启动

```bash
cd ppt-image-tool
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

然后在浏览器打开 http://127.0.0.1:5000 。

## 使用流程

1. **填 Key**：在「API 设置」里粘贴 DeepSeek 和 OpenAI 两个 Key，可「保存到本机」下次自动带出。
2. **写页面**：点「添加页面」，逐页用自然语言描述内容（可拖动排序）。
3. **优化提示词**：点「优化提示词」（或「一键优化全部」），DeepSeek 会把描述改写成专业生图提示词，可手动微调。
4. **生成图片**：点「生成图片」（或「一键生成全部图片」），等待每页出图（约 10-60 秒/页）。个别页面失败时，可点「重新生成失败图片」只重试失败的页。
5. **拼 PPT**：点「生成 PPTX」下载 `.pptx`；点「下载全部图片」打包原图。

## 输出路径自定义

- **图片保存目录**：「API 设置 → 图片保存目录」可自定义页面图、图表、风格母版与参考图的落盘位置（默认 `output/images`，母版/参考图放其 `masters` 子文件夹）。留空即使用默认路径。
- **输出文件保存路径**：单独控制生成的 PPTX 与图片 ZIP 保存到哪里（默认 `output` 目录）。两者互不影响。

## 目录结构

```
ppt-image-tool/
├── app.py            # 后端服务
├── requirements.txt  # 依赖
├── start.bat         # Windows 一键启动
├── static/
│   └── index.html    # 网页前端
└── output/           # 生成的图片与 PPT（自动创建）
    ├── images/       # 页面图 page_*、精确图表 chart_*
    ├── masters/      # 风格母版 master_*、参考图 ref_*
    └── *.pptx
```

## 常见问题

- **图片接口 401**：Key 无效，或 OpenAI 账号未完成 gpt-image-2 的资质验证（Organization Verification）。
- **国内直连 OpenAI 不稳定**：可在「OpenAI API 地址」里填你自用的中转/代理地址（形如 `https://xxx/v1`）。
- **生成太慢 / 太贵**：把「图片质量」调成 `low` 或 `medium`，尺寸用 `1536x1024`。
