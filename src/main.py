import os
import time
import concurrent.futures
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from PIL import Image
import logging
import argparse
import shutil
import re

# --- New Data Structures ---
class ElementData:
    """Represents a generic element extracted from HTML."""
    def __init__(self):
        self.tag_name = None
        self.classes = []
        self.text = None
        self.geom = None
        self.icon_path = None
        self.element_screenshot_path = None # 新增：用于存储元素截图的路径
        self.children = []

class SlideData:
    """Holds all data for a single slide."""
    def __init__(self):
        self.background_image_path = None
        self.elements = []

# --- Constants ---
ICON_CLASSES = {
    'material-icons', 'toc-icon', 'importance-icon', 'limitation-icon', 
    'check-icon', 'partial-icon', 'close-icon', 'feature-icon', 'section-icon', 
    'api-icon', 'config-icon', 'case-icon', 'component-icon', 'mock-icon', 
    'snapshot-icon', 'resource-icon'
}

# --- Layout Constants ---
SLIDE_WIDTH_PX = 1280
SLIDE_RIGHT_MARGIN_PX = 40 # Margin from the right edge of the slide
WIDTH_INCREASE_FACTOR = 1.25 # 25% safety buffer for non-full-width elements
FULL_WIDTH_CLASSES = {'title', 'subtitle', 'section-title', 'p-full-width'}


# --- Selenium and Parsing Logic ---

def parse_color(color_str):
    """解析CSS颜色字符串 (rgb, rgba) 并返回一个元组 (r, g, b, a)。"""
    if not color_str or color_str in ['transparent', 'inherit', 'initial', 'unset']:
        return (0, 0, 0, 0)
    try:
        # 查找字符串中的所有数字
        parts = re.findall(r'[\d.]+', color_str)
        r, g, b = int(parts[0]), int(parts[1]), int(parts[2])
        a = float(parts[3]) if len(parts) > 3 else 1.0
        return (r, g, b, a)
    except (IndexError, ValueError):
        return (0, 0, 0, 0) # 解析错误时默认为透明黑色

def init_driver():
    """Initializes the Selenium WebDriver."""
    options = webdriver.ChromeOptions()
    options.add_argument("--window-size=1280,720")
    options.add_argument("--hide-scrollbars")
    options.add_argument('--headless')  # 添加无头模式
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def take_icon_screenshot(driver, icon_element, temp_dir, slide_index, element_index, slide_element=None):
    """Takes a high-resolution, cropped screenshot of an icon element."""
    # 在截图前先将slide_element向下移动1000px，防止内容干扰
    if slide_element:
        try:
            driver.execute_script("arguments[0].style.transform = 'translateY(1000px)';", slide_element)
            time.sleep(0.1)  # 等待移动完成
        except Exception as e:
            logging.warning(f"无法移动slide元素: {e}")
    
    scale_factor = 5
    js_script = """
        const targetElement = arguments[0];
        const scale = arguments[1];
        const style = window.getComputedStyle(targetElement);
        const originalFontSizeStr = style.getPropertyValue('font-size');
        const originalColor = style.getPropertyValue('color');
        if (!originalFontSizeStr) return null;
        const originalFontSize = parseFloat(originalFontSizeStr);
        const container = document.createElement('div');
        container.id = 'temp-icon-container-for-screenshot';
        container.style.position = 'absolute';
        container.style.left = '0px';
        container.style.top = '0px';
        container.style.zIndex = '9999';
        const clone = targetElement.cloneNode(true);
        clone.style.fontSize = (originalFontSize * scale) + 'px';
        clone.style.color = originalColor;
        clone.style.backgroundColor = 'transparent';
        container.appendChild(clone);
        document.body.appendChild(container);
        return clone;
    """
    scaled_clone_element = driver.execute_script(js_script, icon_element, scale_factor)
    if not scaled_clone_element:
        logging.warning(f"Could not create a scalable clone for icon on slide {slide_index}.")
        # 恢复slide_element位置
        if slide_element:
            try:
                driver.execute_script("arguments[0].style.transform = '';", slide_element)
            except Exception:
                pass
        return None

    time.sleep(0.1)
    print("Paused before icon screenshot. Press enter to continue...")
    # input()
    icon_path = os.path.join(temp_dir, f"slide_{slide_index}_element_{element_index}_icon.png")
    try:
        scaled_clone_element.screenshot(icon_path)
        with Image.open(icon_path) as img:
            bbox = img.getbbox()
            if bbox:
                cropped_img = img.crop(bbox)
                cropped_img.save(icon_path)
        return icon_path
    except Exception as e:
        logging.error(f"Failed to screenshot or crop icon: {e}")
        return None
    finally:
        driver.execute_script("document.getElementById('temp-icon-container-for-screenshot').remove();")
        # 恢复slide_element位置
        if slide_element:
            try:
                driver.execute_script("arguments[0].style.transform = '';", slide_element)
            except Exception as e:
                logging.warning(f"无法恢复slide元素位置: {e}")

def parse_element_recursively(driver, element, temp_dir, slide_index, element_counter, parent_bg_color=None, slide_element=None):
    """Recursively parses an element and its children to extract data."""
    data = ElementData()
    try:
        data.tag_name = element.tag_name
        data.classes = element.get_attribute('class').split()
    except Exception:
        return None

    # Get geometry and style for all elements first.
    try:
        location = element.location
        size = element.size
        if size['width'] == 0 or size['height'] == 0: # Skip invisible elements
             return None
        data.geom = {
            "x": location['x'], "y": location['y'],
            "width": size['width'], "height": size['height'],
            "font-size": element.value_of_css_property('font-size'),
            "color": element.value_of_css_property('color'),
            "font-weight": element.value_of_css_property('font-weight'),
            "text-align": element.value_of_css_property('text-align'),
            'background-color': element.value_of_css_property('background-color')
        }
    except Exception:
        return None # Element is not visible or interactable

    # Check if the element is an icon. Icons are leaf nodes.
    if any(cls in ICON_CLASSES for cls in data.classes):
        data.icon_path = take_icon_screenshot(driver, element, temp_dir, slide_index, element_counter['i'], slide_element)
        return data

    # --- Reordered Logic ---

    # 1. Extract text and find children BEFORE potential DOM modification
    try:
        js_get_text = "return Array.from(arguments[0].childNodes).filter(node => node.nodeType === 3 && node.nodeValue.trim() !== '').map(node => node.nodeValue.trim()).join(' ')"
        text = driver.execute_script(js_get_text, element)
        if text:
            data.text = text
    except Exception as e:
        logging.warning(f"Could not extract text from element: {e}")

    child_elements = element.find_elements(By.XPATH, "./*")

    # 2. Now, handle the background screenshot using the safe clone method
    try:
        bg_color_str = data.geom.get('background-color')
        bg_color = parse_color(bg_color_str)
    except Exception:
        bg_color = (0, 0, 0, 0)

    is_new_background = bg_color[3] > 0 and bg_color != parent_bg_color
    if is_new_background:
        element_path = os.path.join(temp_dir, f"slide_{slide_index}_element_{element_counter['i']}_bg.png")
        try:
            print("Paused before element background screenshot. Press enter to continue...")
            # input()
            
            # 隐藏element下的所有一级子元素
            child_elements_to_hide = element.find_elements(By.XPATH, "./*")
            original_styles = []
            original_texts = []  # 新增：保存原始文本内容
            for child in child_elements_to_hide:
                try:
                    # 保存原始样式
                    original_style = driver.execute_script("return arguments[0].style.display;", child)
                    original_styles.append((child, original_style))
                    # 隐藏元素
                    driver.execute_script("arguments[0].style.display = 'none';", child)
                except Exception as e:
                    logging.warning(f"无法隐藏子元素: {e}")
                    original_styles.append((child, None))
            
            # 新增：将当前元素的直接文本内容替换为等宽度空白字符
            try:
                # 保存原始文本内容并替换为空白字符
                original_text = driver.execute_script("""
                    var element = arguments[0];
                    var textNodes = [];
                    for (var i = 0; i < element.childNodes.length; i++) {
                        if (element.childNodes[i].nodeType === 3) { // TEXT_NODE
                            var originalText = element.childNodes[i].textContent;
                            textNodes.push(originalText);
                            // 将文本替换为相同长度的空白字符（使用全角空格保持宽度）
                            var spaceText = '\u3000'.repeat(originalText.length);
                            element.childNodes[i].textContent = spaceText;
                        }
                    }
                    return textNodes;
                """, element)
                original_texts.append((element, original_text))
            except Exception as e:
                logging.warning(f"无法替换元素文本: {e}")
                original_texts.append((element, []))
            
            # 等待页面重新渲染
            time.sleep(0.5)
            
            # 截图
            element.screenshot(element_path)
            data.element_screenshot_path = element_path
            
            # 恢复所有一级子元素的显示状态
            for child, original_style in original_styles:
                try:
                    if original_style:
                        driver.execute_script("arguments[0].style.display = arguments[1];", child, original_style)
                    else:
                        driver.execute_script("arguments[0].style.display = '';", child)
                except Exception:
                    logging.warning(f"无法恢复子元素显示状态: {e}")
            
            # 新增：恢复原始文本内容
            for elem, text_list in original_texts:
                try:
                    driver.execute_script("""
                        var element = arguments[0];
                        var textList = arguments[1];
                        var textNodeIndex = 0;
                        for (var i = 0; i < element.childNodes.length; i++) {
                            if (element.childNodes[i].nodeType === 3 && textNodeIndex < textList.length) { // TEXT_NODE
                                element.childNodes[i].textContent = textList[textNodeIndex];
                                textNodeIndex++;
                            }
                        }
                    """, elem, text_list)
                except Exception:
                    logging.warning(f"无法恢复元素文本: {e}")
            
            # 等待页面恢复渲染
            time.sleep(0.2)
                    
        except Exception as e:
            # 如果出错，确保恢复所有一级子元素的显示状态和文本内容
            try:
                for child, original_style in original_styles:
                    try:
                        if original_style:
                            driver.execute_script("arguments[0].style.display = arguments[1];", child, original_style)
                        else:
                            driver.execute_script("arguments[0].style.display = '';", child)
                    except Exception:
                        pass
                # 恢复文本内容
                for elem, text_list in original_texts:
                    try:
                        driver.execute_script("""
                            var element = arguments[0];
                            var textList = arguments[1];
                            var textNodeIndex = 0;
                            for (var i = 0; i < element.childNodes.length; i++) {
                                if (element.childNodes[i].nodeType === 3 && textNodeIndex < textList.length) {
                                    element.childNodes[i].textContent = textList[textNodeIndex];
                                    textNodeIndex++;
                                }
                            }
                        """, elem, text_list)
                    except Exception:
                        pass
            except Exception:
                pass
            logging.warning(f"Could not take element background screenshot for slide {slide_index}: {e}")

    # 3. Recursively parse the children we found earlier
    current_bg_for_children = bg_color if is_new_background else parent_bg_color
    for child_element in child_elements:
        element_counter['i'] += 1
        child_data = parse_element_recursively(driver, child_element, temp_dir, slide_index, element_counter, parent_bg_color=current_bg_for_children, slide_element=slide_element)
        if child_data:
            data.children.append(child_data)

    # Pruning
    if not data.text and not data.icon_path and not data.element_screenshot_path and not data.children:
        return None

    return data

def wait_for_material_icons(driver, timeout=10):
    """等待Material Icons字体加载完成"""
    js_script = """
        return new Promise((resolve) => {
            if (document.fonts && document.fonts.ready) {
                document.fonts.ready.then(() => {
                    // 额外等待一下确保material-icons完全加载
                    setTimeout(() => resolve(true), 500);
                });
            } else {
                // 如果不支持document.fonts API，则等待固定时间
                setTimeout(() => resolve(true), 2000);
            }
        });
    """
    
    try:
        driver.set_script_timeout(timeout)
        driver.execute_async_script(js_script)
        logging.info("Material Icons字体加载完成")
    except Exception as e:
        logging.warning(f"等待Material Icons加载时出错: {e}，继续执行")

def extract_data_from_html(driver, file_path, temp_dir):
    """Extracts structured data from all slides in the HTML file using a pre-initialized driver."""
    driver.get(f"file:///{os.path.abspath(file_path)}")
    time.sleep(2) # Allow time for rendering
    
    # 等待Material Icons字体加载完成
    wait_for_material_icons(driver)

    slides_data = []
    slide_elements = driver.find_elements(By.CSS_SELECTOR, ".slide")
    logging.info(f"Found {len(slide_elements)} slides in {os.path.basename(file_path)}.")

    for i, slide_element in enumerate(slide_elements):
        logging.info(f"Processing slide {i+1}...")
        slide_data = SlideData()
        element_counter = {'i': 0} # Use a mutable dict for a counter

        # 1. Take background screenshot
        screenshot_path = os.path.join(temp_dir, f"slide_{i}_bg.png")
        try:
            print("Paused before slide background screenshot. Press enter to continue...")
            # input()
            
            # 隐藏slide_element下的所有子元素
            child_elements_to_hide = slide_element.find_elements(By.XPATH, ".//*")
            original_styles = []
            for child in child_elements_to_hide:
                try:
                    # 保存原始样式
                    original_style = driver.execute_script("return arguments[0].style.display;", child)
                    original_styles.append((child, original_style))
                    # 隐藏元素
                    driver.execute_script("arguments[0].style.display = 'none';", child)
                except Exception as e:
                    logging.warning(f"无法隐藏子元素: {e}")
                    original_styles.append((child, None))
            
            # 截图
            slide_element.screenshot(screenshot_path)
            slide_data.background_image_path = screenshot_path
            logging.info(f"Took background screenshot for slide {i+1}.")
            
            # 恢复所有子元素的显示状态
            for child, original_style in original_styles:
                try:
                    if original_style:
                        driver.execute_script("arguments[0].style.display = arguments[1];", child, original_style)
                    else:
                        driver.execute_script("arguments[0].style.display = '';", child)
                except Exception as e:
                    logging.warning(f"无法恢复子元素显示状态: {e}")
            
        except Exception as e:
            # 如果出错，确保恢复所有子元素的显示状态
            try:
                for child, original_style in original_styles:
                    try:
                        if original_style:
                            driver.execute_script("arguments[0].style.display = arguments[1];", child, original_style)
                        else:
                            driver.execute_script("arguments[0].style.display = '';", child)
                    except Exception:
                        pass
            except Exception:
                pass
            logging.error(f"Could not take background screenshot for slide {i+1}: {e}")

        # 2. Recursively parse content
        try:
            try:
                content_element = slide_element.find_element(By.CSS_SELECTOR, ".slide-content")
            except Exception:
                # Fallback to using the slide itself if no .slide-content is found
                logging.warning(f"No .slide-content found in slide {i+1}. Parsing from .slide root.")
                content_element = slide_element

            child_elements = content_element.find_elements(By.XPATH, "./*")
            for child in child_elements:
                element_counter['i'] += 1
                element_data = parse_element_recursively(driver, child, temp_dir, i, element_counter, slide_element=slide_element)
                if element_data:
                    slide_data.elements.append(element_data)
        except Exception as e:
            logging.error(f"Error processing content for slide {i+1}: {e}")

        slides_data.append(slide_data)

    return slides_data

# --- PowerPoint Generation Logic ---

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs

def px_to_emu(px):
    if isinstance(px, (int, float)):
        return int(px * 9525)
    return 0

def add_slide_with_background(prs, image_path):
    slide_layout = prs.slide_layouts[6] # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    if image_path and os.path.exists(image_path):
        try:
            slide.shapes.add_picture(image_path, 0, 0, width=prs.slide_width, height=prs.slide_height)
        except Exception as e:
            logging.error(f"Failed to add background image {image_path}: {e}")
    return slide

def add_image(slide, image_path, geom):
    if not (image_path and os.path.exists(image_path) and geom): return
    try:
        # Use the geometry of the icon itself for positioning and sizing
        x_emu = px_to_emu(geom['x'])
        y_emu = px_to_emu(geom['y'])
        width_emu = px_to_emu(geom['width'])
        height_emu = px_to_emu(geom['height'])
        
        # Ensure minimum size to avoid errors with zero-sized elements
        if width_emu == 0 or height_emu == 0:
            default_size = px_to_emu(20) # Default size for tiny/invisible icons
            width_emu = width_emu or default_size
            height_emu = height_emu or default_size

        slide.shapes.add_picture(image_path, x_emu, y_emu, width=width_emu, height=height_emu)
    except Exception as e:
         logging.error(f"Failed to add icon image {image_path}: {e}")

def add_textbox(slide, data, slide_width_px):
    """Adds a textbox to the slide with dynamic width adjustment."""
    if not data or not data.text or not data.geom:
        return

    geom = data.geom.copy()  # Work on a copy to avoid side effects
    text = data.text

    # --- Dynamic width adjustment logic ---
    is_full_width = any(cls in FULL_WIDTH_CLASSES for cls in data.classes)

    if is_full_width:
        # For titles, subtitles, etc., extend the width towards the slide margin
        available_width = slide_width_px - geom['x'] - SLIDE_RIGHT_MARGIN_PX
        # Use the larger of the original or the calculated width
        geom['width'] = max(geom['width'], available_width)
    else:
        # For other elements, add a safety buffer to prevent wrapping
        geom['width'] *= WIDTH_INCREASE_FACTOR
    # --- End of adjustment ---

    try:
        x_emu = px_to_emu(geom['x'])
        y_emu = px_to_emu(geom['y'])
        width_emu = px_to_emu(geom['width'])
        height_emu = px_to_emu(geom['height'])

        textbox = slide.shapes.add_textbox(x_emu, y_emu, width_emu, height_emu)
        text_frame = textbox.text_frame
        text_frame.word_wrap = True
        p = text_frame.paragraphs[0]
        run = p.add_run()
        run.text = text

        font = run.font
        if 'font-size' in geom:
            try:
                font_size_px = float(str(geom['font-size']).replace("px", ""))
                scale_factor = 0.75  # Adjusted scale factor for better pptx rendering
                font.size = Pt(int(font_size_px * scale_factor))
            except (ValueError, TypeError):
                pass
        if 'color' in geom:
            try:
                color_str = geom['color'].replace("rgba(", "").replace("rgb(", "").replace(")", "")
                parts = [p.strip() for p in color_str.split(",")]
                r, g, b = int(parts[0]), int(parts[1]), int(parts[2])
                font.color.rgb = RGBColor(r, g, b)
            except Exception:
                pass
        if 'font-weight' in geom and (str(geom['font-weight']) == 'bold' or (str(geom['font-weight']).isnumeric() and int(geom['font-weight']) >= 700)):
            font.bold = True
    except Exception as e:
        logging.error(f"Failed to add textbox with text '{text}': {e}")


def add_elements_to_slide(slide, elements, slide_width_px):
    """Recursively adds elements from ElementData to the slide."""
    for data in elements:
        # 1. 如果有背景截图，先添加背景
        if data.element_screenshot_path:
            add_image(slide, data.element_screenshot_path, data.geom)

        # 2. 添加元素本身的内容（图标或文本）
        if data.icon_path:
            add_image(slide, data.icon_path, data.geom)
        elif data.text:
            add_textbox(slide, data, slide_width_px)

        # 3. 递归添加子元素
        if data.children:
            add_elements_to_slide(slide, data.children, slide_width_px)


# --- Main Execution Logic ---

def process_files_worker(task_info):
    """
    Worker function for a thread. Initializes a single WebDriver instance
    and processes a chunk of HTML files.
    """
    files_chunk, input_dir, output_dir, temp_dir = task_info
    
    if not files_chunk:
        return # Nothing to do for this worker

    logging.info(f"Worker starting, assigned {len(files_chunk)} files.")
    driver = None
    try:
        driver = init_driver()
        for html_file in files_chunk:
            try:
                logging.info(f"--- Processing file: {html_file} ---")

                # Each thread gets its own temp sub-directory to avoid file collisions
                thread_temp_dir = os.path.join(temp_dir, re.sub(r'[^a-zA-Z0-9.-]', '_', html_file))
                if not os.path.exists(thread_temp_dir):
                    os.makedirs(thread_temp_dir)

                # Create a new presentation for each file
                prs = create_presentation()

                file_path = os.path.join(input_dir, html_file)
                # Call the modified extraction function
                all_slides_data = extract_data_from_html(driver, file_path, thread_temp_dir)

                for slide_data in all_slides_data:
                    slide = add_slide_with_background(prs, slide_data.background_image_path)
                    add_elements_to_slide(slide, slide_data.elements, SLIDE_WIDTH_PX)

                # Save the presentation
                base_name = os.path.splitext(html_file)[0]
                output_path = os.path.join(output_dir, f"{base_name}.pptx")
                prs.save(output_path)
                logging.info(f"Successfully created {output_path}")
            except Exception as e:
                # Log exceptions for a single file, but continue with the rest of the chunk
                logging.error(f"Failed to process {html_file}: {e}", exc_info=True)
    except Exception as e:
        # Log exceptions for the entire worker (e.g., driver init failure)
        logging.critical(f"Worker failed catastrophically: {e}", exc_info=True)
    finally:
        if driver:
            driver.quit()
            logging.info(f"Worker finished, WebDriver has been shut down.")

def main():
    parser = argparse.ArgumentParser(description='Convert HTML file(s) to PPTX presentations.')
    parser.add_argument('--input_path', type=str, required=True, help='Path to an input HTML file or a directory containing HTML files.')
    parser.add_argument('--output_dir', type=str, required=True, help='Output directory for the generated PPTX files.')
    parser.add_argument('--workers', type=int, default=os.cpu_count(), help='Number of parallel threads to use for conversion. Defaults to the number of CPU cores.')
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(threadName)s - %(levelname)s - %(message)s')
    
    input_path = args.input_path
    output_dir = args.output_dir

    # Setup temp directory
    base_dir = os.path.dirname(os.path.abspath(__file__))
    temp_dir = os.path.join(base_dir, "..", "temp")
    # if os.path.exists(temp_dir):
    #     shutil.rmtree(temp_dir)
    os.makedirs(temp_dir, exist_ok=True)

    # Ensure output directory exists
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # --- Determine input files ---
    html_files = []
    input_base_dir = None

    if not os.path.exists(input_path):
        logging.error(f"Input path not found: {input_path}")
        return

    if os.path.isdir(input_path):
        input_base_dir = input_path
        try:
            html_files = [f for f in os.listdir(input_base_dir) if f.endswith('.html')]
            # 修改排序逻辑，更精确地提取文件名中的数字
            def extract_number(filename):
                match = re.search(r'file_(\d+)\.html', filename)
                return int(match.group(1)) if match else 0
            
            html_files.sort(key=extract_number)
            logging.info(f"Found and sorted {len(html_files)} HTML files in directory '{input_base_dir}'.")
        except Exception as e:
            logging.error(f"Error reading input directory '{input_base_dir}': {e}")
            return
    elif os.path.isfile(input_path):
        if input_path.endswith('.html'):
            input_base_dir = os.path.dirname(input_path)
            html_files = [os.path.basename(input_path)]
            logging.info(f"Found single HTML file: '{input_path}'")
        else:
            logging.error(f"Input file is not an HTML file: '{input_path}'")
            return
    else:
        logging.error(f"Input path is not a valid file or directory: '{input_path}'")
        return
    
    logging.info(f"Starting conversion for {len(html_files)} file(s) to output directory '{output_dir}' with {args.workers} workers.")

    # Distribute files among workers
    num_workers = min(args.workers, len(html_files))
    if num_workers == 0:
        logging.info("No HTML files to process.")
        return
        
    # Create chunks of files for each worker
    file_chunks = [[] for _ in range(num_workers)]
    for i, html_file in enumerate(html_files):
        file_chunks[i % num_workers].append(html_file)

    tasks = [(chunk, input_base_dir, output_dir, temp_dir) for chunk in file_chunks]

    # Process files in parallel
    with concurrent.futures.ThreadPoolExecutor(max_workers=num_workers, thread_name_prefix='Converter') as executor:
        executor.map(process_files_worker, tasks)

    # Clean up temp files
    # if os.path.exists(temp_dir):
    #     shutil.rmtree(temp_dir)
    #     logging.info("Temporary files cleaned up.")

if __name__ == "__main__":
    main()