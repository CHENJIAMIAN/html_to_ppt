# 技术栈

## 核心技术
- **Python 3.x** - 主要编程语言
- **Selenium WebDriver** - 浏览器自动化，用于HTML渲染和截图捕获
- **python-pptx** - PowerPoint文件生成和操作
- **PIL (Pillow)** - 图像处理和操作
- **Chrome/Chromium** - 无头浏览器用于渲染

## 关键库
- `selenium` - 网页自动化和元素交互
- `webdriver-manager` - 自动ChromeDriver管理
- `python-pptx` - PowerPoint演示文稿创建
- `PIL/Pillow` - 图像裁剪和处理
- `concurrent.futures` - 多线程支持
- `logging` - 应用程序日志记录
- `argparse` - 命令行界面

## 构建和运行命令

### 环境设置
```bash
pip install -r requirements.txt
```

### 基本用法
```bash
python src/main.py --input_path input/ --output_dir output/
```

### 高级用法
```bash
# 指定工作线程数量
python src/main.py --input_path input/file.html --output_dir output/ --workers 4

# 处理单个文件
python src/main.py --input_path input/presentation.html --output_dir output/
```

## 架构模式
- **多线程处理** - 每个工作线程拥有自己的WebDriver实例
- **递归元素解析** - DOM树遍历，保持父子关系
- **基于截图的渲染** - 将视觉元素捕获为图像以实现精确布局
- **模块化设计** - 分离解析、渲染和PowerPoint生成的关注点