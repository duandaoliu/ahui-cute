# 项目说明

这是一个用于脚本测试和自动化工具的实验目录，非正式项目仓库。

## 用户偏好

- 操作系统：Windows 11，常用 bash (Git Bash) 作为终端
- 语言：中文为主
- 编程：偏好 Python 脚本实现自动化
- 办公软件：通过 COM 接口自动化操控 Office（而非后台生成文件），喜欢看到实时操作过程

## 可用环境

- **Python**: Python 3.7.1 (Miniconda3)，另有 Python 3.13 备用
- **关键库**: pywin32 (COM 自动化), Pillow (图像处理)
- **路径**: 项目脚本放在 `C:\Users\33164\Desktop\todo\`
- **桌面**: 文件保存默认到桌面 `C:\Users\33164\Desktop\`

## 视觉识别

- 使用豆包模型 `doubao-seed-2-0-lite-260428` 做图片识别
- API: 火山方舟 Ark，endpoint: `https://ark.cn-beijing.volces.com/api/v3/chat/completions`
- **图片**: `vision.py`，用法: `python vision.py <图片路径> [可选问题]`
- **PDF**: `read_pdf.py`，用法: `python read_pdf.py <PDF路径> [页码] [可选问题]` — 自动将PDF页面转图片后调用豆包识别
- API Key: ark-74990dfa-9292-47a5-9ade-1ac81567a98c-052c0

## 自动化规范

### PowerPoint
- 使用 `win32com.client` COM 接口直接操控 PowerPoint 应用程序
- PowerPoint 设置为可见 (`ppt.Visible = True`)，用户可实时观看生成过程
- 保存后不关闭 PowerPoint，让用户自行编辑
- 注意 Windows 控制台编码问题，避免在 print 中使用 emoji

### SolidWorks
- SolidWorks 同样支持 COM API，可通过 win32com 或 VBA 自动化
- 用户有意向用脚本自动建模，待具体需求

## 文件组织

```
todo/
├── create_ppt.py    # PowerPoint 自动生成 PPT 脚本
├── CLAUDE.md        # 本项目说明文件
```

## 沟通风格

- 直接动手，少解释，用代码说话
- 运行脚本时展示过程，让用户看到效果
- 遇到问题快速修复，不纠结
