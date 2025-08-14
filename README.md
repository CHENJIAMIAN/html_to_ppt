# HTML转PowerPoint转换器

一个将HTML演示文稿高保真转换为PowerPoint幻灯片的Python工具。

## 功能特性

- 🚀 **多线程并行处理** - 支持批量转换多个HTML文件
- 🎯 **精确布局保持** - 保持原始HTML的布局和样式
- 🖼️ **高分辨率图标** - 自动截图并优化图标显示
- 🎨 **样式完整支持** - 支持背景色、圆角、阴影等CSS样式
- 📱 **Material Icons** - 完整支持Material Icons字体
- ⚡ **自动化处理** - 无需手动调整，一键转换

## 技术栈

- **Python 3.x** - 主要编程语言
- **Selenium WebDriver** - 浏览器自动化和HTML渲染
- **python-pptx** - PowerPoint文件生成
- **PIL (Pillow)** - 图像处理和优化
- **Chrome/Chromium** - 无头浏览器渲染引擎

## 安装要求

### 系统要求
- Python 3.7+
- Chrome浏览器（用于HTML渲染）

### 安装依赖
```bash
pip install -r requirements.txt
```

## 使用方法

### 基本用法
```bash
# 转换单个HTML文件
python src/main.py --input_path input/presentation.html --output_dir output/

# 批量转换整个目录
python src/main.py --input_path input/ --output_dir output/
```

### 高级用法
```bash
# 指定工作线程数量（默认为CPU核心数）
python src/main.py --input_path input/ --output_dir output/ --workers 4

# 启用详细日志输出
python src/main.py --input_path input/ --output_dir output/ --verbose
```

## 项目结构

```
├── src/
│   └── main.py              # 主应用程序入口点
├── input/                   # 待转换的HTML文件
├── output/                  # 生成的PowerPoint文件
├── temp/                    # 临时文件（自动清理）
├── requirements.txt         # Python依赖项
├── README.md               # 项目说明文档
└── .gitignore              # Git忽略规则
```

## HTML格式要求

转换器支持标准的HTML演示文稿格式，要求：

1. **幻灯片结构**：使用 `.slide` 类定义每张幻灯片
2. **标题区域**：使用 `.slide-header` 类定义标题区域
3. **内容区域**：使用 `.slide-content` 类定义内容区域
4. **图标支持**：支持Material Icons和自定义图标类

### 示例HTML结构
```html
<div class="slide">
    <div class="slide-header">
        <h1 class="title">幻灯片标题</h1>
        <p class="subtitle">副标题</p>
    </div>
    <div class="slide-content">
        <!-- 幻灯片内容 -->
    </div>
</div>
```

## 支持的CSS特性

- ✅ 文本样式（字体、颜色、大小、粗细）
- ✅ 背景色（包括透明度）
- ✅ 圆角边框
- ✅ 阴影效果
- ✅ 精确定位和尺寸
- ✅ Material Icons图标

## 常见问题

### Q: 转换后的PowerPoint文件中图标显示异常？
A: 确保HTML文件中正确引入了Material Icons字体，转换器会自动等待字体加载完成。

### Q: 如何提高转换速度？
A: 可以通过 `--workers` 参数增加并行处理的线程数量，建议设置为CPU核心数。

### Q: 支持哪些浏览器？
A: 目前仅支持Chrome/Chromium浏览器，转换器会自动下载和管理ChromeDriver。

## 开发说明

### 核心组件
- **ElementData** - HTML元素数据容器
- **SlideData** - 幻灯片数据容器
- **parse_element_recursively()** - 递归DOM解析
- **create_presentation()** - PowerPoint生成

### 多线程架构
- 每个工作线程拥有独立的WebDriver实例
- 线程安全的临时文件管理
- 自动资源清理和错误恢复

## 许可证

MIT License - 详见 [LICENSE](LICENSE) 文件

## 贡献

欢迎提交Issue和Pull Request来改进这个项目！

## 更新日志

### v1.0.0
- 初始版本发布
- 支持基本的HTML到PowerPoint转换
- 多线程并行处理
- Material Icons支持