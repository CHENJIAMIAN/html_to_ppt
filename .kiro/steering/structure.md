# 项目结构

## 目录布局
```
├── src/
│   └── main.py              # 主应用程序入口点
├── input/                   # 待转换的HTML文件
├── output/                  # 生成的PowerPoint文件
├── temp/                    # 临时文件（截图、处理产物）
├── .kiro/
│   └── steering/           # AI助手指导文档
├── requirements.txt        # Python依赖项
├── todolist.md            # 已知问题和改进项
└── .gitignore             # Git忽略模式
```

## 代码组织

### 主模块 (`src/main.py`)
- **数据类**: `ElementData`, `SlideData` - 核心数据结构
- **配置**: 布局、样式和处理的常量
- **核心逻辑**: 
  - `parse_element_recursively()` - DOM遍历和数据提取
  - `extract_data_from_html()` - 主要HTML处理管道
  - `create_presentation()` - PowerPoint生成
- **工作函数**: `process_files_worker()` - 多线程处理
- **工具函数**: 颜色解析、坐标转换、截图处理

## 文件命名约定
- 输入文件: `input/` 目录中的 `*.html`
- 输出文件: `output/` 目录中的 `{basename}.pptx`
- 临时文件: `temp/` 子目录中的 `slide_{index}_{type}.png`
- 线程隔离: 每个工作线程使用 `temp/{清理后的文件名}/` 子目录

## 关键设计模式
- **工厂模式**: WebDriver初始化，配置一致
- **递归处理**: 元素树遍历，保持父子关系
- **工作池**: 线程安全的并行处理，资源隔离
- **数据传输对象**: 元素和幻灯片信息的结构化数据容器

## 开发指南
- 所有临时文件都会自动管理和清理
- 每个线程使用隔离的WebDriver实例和临时目录
- 截图文件使用描述性命名便于调试
- 日志记录具有线程感知能力，能够正确识别