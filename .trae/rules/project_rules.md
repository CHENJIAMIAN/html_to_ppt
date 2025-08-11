# Todolist: 构建基于约定格式的HTML到PPTX转换器

本项目旨在创建一个Python脚本，该脚本能读取遵循预定义模板的HTML文件，并将其高质量地转换为可编辑的PowerPoint (.pptx) 演示文稿。

## Phase 0: 项目初始化与环境搭建

此阶段的目标是建立一个稳定、可复现的开发环境。

-   [x] **1. 创建项目目录结构**
    -   [x] 创建主目录, e.g., `html_to_pptx_converter`
    -   [x] 在主目录下创建子目录:
        -   `src/`: 存放核心Python脚本。
        -   `input/`: 存放待转换的HTML文件。
        -   `output/`: 存放生成的PPTX文件。
        -   `temp/`: 存放临时的截图文件。

-   [x] **2. 设置Python虚拟环境**
    -   [x] 在项目根目录打开终端。
    -   [x] 执行 `python -m venv venv` 创建虚拟环境。
    -   [x] 激活虚拟环境 (Windows: `venv\Scripts\activate`, macOS/Linux: `source venv/bin/activate`)。

-   [x] **3. 安装必要的Python库**
    -   [x] 创建 `requirements.txt` 文件。
    -   [x] 将以下依赖项添加到文件中：
        ```
        python-pptx
        selenium
        webdriver-manager
        beautifulsoup4
        lxml 
        pillow 
        ```
    -   [x] 在激活的虚拟环境中运行 `pip install -r requirements.txt`。

-   [x] **4. 创建主脚本文件**
    -   [x] 在 `src/` 目录下创建 `main.py` 文件。
    -   [x] 在脚本顶部导入所有必要的库。

## Phase 1: HTML解析与数据提取引擎（输入端）

此阶段的核心是利用Selenium驱动真实浏览器来加载HTML，并精确提取每个“幻灯片”的结构化数据。

-   [x] **1. 初始化Selenium Web Driver**
    -   [x] 在 `main.py` 中编写一个函数，用于设置和启动Chrome浏览器。
    -   [x] **关键点**: 设置浏览器窗口大小与HTML中的 `.slide` 尺寸一致 (`1280x720`)，以确保坐标和尺寸的准确性。
    -   [x] 使用`webdriver-manager`自动管理ChromeDriver。

-   [x] **2. 定义数据结构**
    -   [x] 设计一个Python类或字典结构来存储从单个HTML幻灯片中提取的数据。例如，一个 `SlideData` 类。
    -   [x] 该结构应包含字段：`background_image_path`, `title_text`, `title_geom`, `subtitle_text`, `subtitle_geom`, `keyword_items` (一个列表) 等。`geom` (geometry) 应包含 `(x, y, width, height)` 和字体信息。

-   [x] **3. 实现HTML数据提取逻辑**
    -   [x] 编写一个主函数 `extract_data_from_html(file_path)`。
    -   [x] 使用Selenium加载HTML文件: `driver.get("file:///path/to/your/input.html")`。
    -   [x] 找到所有的幻灯片元素: `driver.find_elements(By.CSS_SELECTOR, ".slide")`。
    -   [x] **循环处理每个幻灯片元素**:
        -   [x] **提取背景**: 定位到 `.slide-background` 元素，使用Selenium的截图功能 (`element.screenshot()`) 对其截图，保存到 `temp/` 目录，并将路径存入`SlideData`对象。
        -   [x] **提取标题**: 定位到 `.title`。获取其 `text`, `location` (x,y坐标), `size` (width,height)，以及通过 `value_of_css_property()` 获取 `font-size`, `color`, `font-weight` 等计算后样式。将这些信息存入`SlideData`对象。
        -   [x] **提取副标题**: 对 `.subtitle` 重复上述过程。
        -   [x] **提取关键词模块 (循环)**:
            -   [x] 定位所有的 `.keyword-item`。
            -   [x] 循环遍历每个`item`。
            -   [x] **提取图标**: 定位到 `<i>` 元素，对其单独截图，保存到`temp/`并记录路径。
            -   [x] **提取关键词标题**: 对 `.keyword-title` 获取文本和几何/样式信息。
            -   [x] **提取关键词描述**: 对 `.keyword-desc` 获取文本和几何/样式信息。
    -   [x] 函数最终返回一个包含所有 `SlideData` 对象的列表。

## Phase 2: PowerPoint生成引擎（输出端）

此阶段负责创建PPTX文件，并提供添加各种元素（背景、文本框、图片）的原子功能。

-   [x] **1. 创建PPTX文件和设置尺寸**
    -   [x] 编写一个函数 `create_presentation()`。
    -   [x] 使用 `pptx.Presentation()` 创建演示文稿。
    -   [x] **关键点**: 设置幻灯片尺寸以匹配16:9的宽高比。`prs.slide_width = Inches(13.333)` 和 `prs.slide_height = Inches(7.5)`（对应1280x720像素）。

-   [x] **2. 实现单位转换工具函数**
    -   [x] **关键点**: `python-pptx` 使用EMUs (English Metric Units)，而Selenium提供像素(px)。必须编写转换函数。
    -   [x] 创建 `px_to_emu(px)` 函数 (1 pixel = 9525 EMUs，在96 DPI下)。
    -   [x] 创建 `pt_to_pt(pt)` 函数 (用于字号，1 pt = 1 pt，通常无需转换，但要明确)。

-   [x] **3. 封装PPT元素添加函数**
    -   [x] 创建 `add_slide_with_background(prs, image_path)`: 添加新幻灯片，并将指定图片设置为幻灯片背景。
    -   [x] 创建 `add_image(slide, image_path, x_emu, y_emu, width_emu)`: 在幻灯片的指定位置和尺寸添加图片（用于图标）。
    -   [x] 创建 `add_textbox(slide, text, x_emu, y_emu, width_emu, height_emu, font_details)`: 在指定位置创建文本框，并根据 `font_details`（一个包含字体名称、大小、颜色、粗细、对齐方式的字典）设置文本格式。

## Phase 3: 逻辑编排与执行

此阶段将输入端和输出端连接起来，完成从数据到演示文稿的完整转换流程。

-   [x] **1. 编写主执行逻辑 `main()`**
    -   [x] 定义输入HTML文件路径和输出PPTX文件路径。
    -   [x] **调用Phase 1**: `all_slides_data = extract_data_from_html(input_path)`。
    -   [x] **调用Phase 2**: `prs = create_presentation()`。
    -   [x] **循环遍历 `all_slides_data` 列表**:
        -   [x] 对于每个 `slide_data` 对象:
        -   [x] 调用 `add_slide_with_background()` 添加带背景的新幻灯片。
        -   [x] 调用 `add_textbox()` 添加标题，传入从`slide_data`中获取的文本、几何和样式信息（记得单位转换）。
        -   [x] 调用 `add_textbox()` 添加副标题。
        -   [x] **内循环遍历 `slide_data.keyword_items`**:
            -   [x] 调用 `add_image()` 添加图标。
            -   [x] 调用 `add_textbox()` 添加关键词标题。
            -   [x] 调用 `add_textbox()` 添加关键词描述。
    -   [x] **保存文件**: `prs.save(output_path)`。

-   [x] **2. 清理临时文件**
    -   [x] 在脚本执行成功后，添加代码删除 `temp/` 目录下的所有临时截图。

## Phase 4: 健壮性与用户体验优化

此阶段的目标是让脚本更可靠、更易用。

-   [x] **1. 添加错误处理**
    -   [x] 在文件操作、元素查找等关键步骤使用 `try...except` 块。
    -   [x] 如果HTML中缺少某个约定元素（如没有`.subtitle`），程序应能优雅地跳过而不是崩溃。

-   [x] **2. 增加日志输出**
    -   [x] 使用 `logging` 模块。
    -   [x] 输出关键步骤信息，如 "正在初始化浏览器...", "发现 5 个幻灯片...", "正在处理幻灯片 3...", "PPTX 文件已保存到 output/presentation.pptx"。

-   [x] **3. 使用命令行参数**
    -   [x] 集成 `argparse` 模块。
    -   [x] 允许用户通过命令行指定输入文件和输出文件，而不是在代码中硬编码。例如: `python src/main.py --input input/mypres.html --output output/mypres.pptx`。

-   [x] **4. 编写项目文档**
    -   [x] 创建一个 `README.md` 文件。
    -   [x] 说明项目的用途、如何安装依赖、如何运行脚本，并详细描述必须遵守的HTML模板格式。
