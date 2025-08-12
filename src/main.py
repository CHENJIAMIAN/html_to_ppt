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

class SlideData:
    def __init__(self):
        self.background_image_path = None
        self.title_text = None
        self.title_geom = None
        self.subtitle_text = None
        self.subtitle_geom = None
        self.keyword_items = []

# --- New Data Structures ---
class ElementData:
    """Represents a generic element extracted from HTML."""
    def __init__(self):
        self.tag_name = None
        self.classes = []
        self.text = None
        self.geom = None
        self.icon_path = None
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

def init_driver():
    """Initializes the Selenium WebDriver."""
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--window-size=1280,720")
    options.add_argument("--hide-scrollbars")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def take_icon_screenshot(driver, icon_element, temp_dir, slide_index, element_index):
    """Takes a high-resolution, cropped screenshot of an icon element."""
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
        return None

    time.sleep(0.1)
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

def parse_element_recursively(driver, element, temp_dir, slide_index, element_counter):
    """Recursively parses an element and its children to extract data."""
    data = ElementData()
    try:
        data.tag_name = element.tag_name
        data.classes = element.get_attribute('class').split()
    except Exception:
        # This can happen if the element becomes stale
        return None

    # Get geometry and style for all elements
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
        }
    except Exception:
        return None # Element is not visible or interactable

    # Check if the element is an icon
    if any(cls in ICON_CLASSES for cls in data.classes):
        data.icon_path = take_icon_screenshot(driver, element, temp_dir, slide_index, element_counter['i'])
        # Icons are considered leaf nodes, don't process children or text
        return data

    # Get geometry and style for all elements
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
        }
    except Exception:
        return None # Element is not visible or interactable

    # Extract text that belongs directly to this element, not its children
    try:
        js_get_text = "return Array.from(arguments[0].childNodes).filter(node => node.nodeType === 3 && node.nodeValue.trim() !== '').map(node => node.nodeValue.trim()).join(' ')"
        text = driver.execute_script(js_get_text, element)
        if text:
            data.text = text
    except Exception as e:
        logging.warning(f"Could not extract text from element: {e}")

    # Recursively parse children
    child_elements = element.find_elements(By.XPATH, "./*")
    for child_element in child_elements:
        element_counter['i'] += 1
        child_data = parse_element_recursively(driver, child_element, temp_dir, slide_index, element_counter)
        if child_data:
            data.children.append(child_data)
            
    # Pruning: if an element has no text, no icon, and no children with content, it's not useful
    if not data.text and not data.icon_path and not data.children:
        return None

    return data

def extract_data_from_html(driver, file_path, temp_dir):
    """Extracts structured data from all slides in the HTML file using a pre-initialized driver."""
    driver.get(f"file:///{os.path.abspath(file_path)}")
    time.sleep(2) # Allow time for rendering

    slides_data = []
    slide_elements = driver.find_elements(By.CSS_SELECTOR, ".slide")
    logging.info(f"Found {len(slide_elements)} slides in {os.path.basename(file_path)}.")

    for i, slide_element in enumerate(slide_elements):
        logging.info(f"Processing slide {i+1}...")
        slide_data = SlideData()
        element_counter = {'i': 0} # Use a mutable dict for a counter

        # 1. Take background screenshot by temporarily hiding text
        original_html = driver.execute_script("return arguments[0].innerHTML;", slide_element)
        try:
            # This JS finds all non-empty text nodes and wraps them in a temporary, hidden span.
            # This is safer than string manipulation of the innerHTML.
            driver.execute_script("""
                const element = arguments[0];
                const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT, null, false);
                let node;
                while(node = walker.nextNode()) {
                    if (node.nodeValue.trim() !== '') {
                        const span = document.createElement('span');
                        // Using visibility:hidden keeps the layout intact, unlike display:none
                        span.style.visibility = 'hidden';
                        node.parentNode.insertBefore(span, node);
                        span.appendChild(node);
                    }
                }
            """, slide_element)
            
            time.sleep(0.2) # Allow re-render

            screenshot_path = os.path.join(temp_dir, f"slide_{i}_bg.png")
            slide_element.screenshot(screenshot_path)
            slide_data.background_image_path = screenshot_path
            logging.info(f"Took background screenshot for slide {i+1} with text hidden.")

        except Exception as e:
            logging.warning(f"Could not take background screenshot for slide {i+1} with text hidden: {e}. A full screenshot will be attempted after restoring HTML.")
            # Fallback screenshot will happen after HTML is restored.
        finally:
            # Restore original HTML to ensure all elements are available for parsing.
            driver.execute_script("arguments[0].innerHTML = arguments[1];", slide_element, original_html)
            time.sleep(0.1) # Allow re-render
            
            # If the background path is still not set, it means the try block failed.
            # Take a screenshot of the restored slide as a fallback.
            if not slide_data.background_image_path:
                try:
                    logging.info(f"Taking fallback full screenshot for slide {i+1}.")
                    screenshot_path = os.path.join(temp_dir, f"slide_{i}_bg.png")
                    slide_element.screenshot(screenshot_path)
                    slide_data.background_image_path = screenshot_path
                except Exception as e_fallback:
                    logging.error(f"Could not take any background screenshot for slide {i+1}: {e_fallback}")

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
                element_data = parse_element_recursively(driver, child, temp_dir, i, element_counter)
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
                scale_factor = 0.85  # Adjusted scale factor for better pptx rendering
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
        # Add the element itself
        if data.icon_path:
            add_image(slide, data.icon_path, data.geom)
        elif data.text:
            add_textbox(slide, data, slide_width_px)

        # Add its children
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
            html_files.sort(key=lambda f: int(re.sub(r'[^0-9]', '', f) or 0))
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