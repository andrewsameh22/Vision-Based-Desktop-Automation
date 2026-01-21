import cv2
import numpy as np
import pyautogui
import time
import os

class VisualGrounding:
    def __init__(self, template_pattern="notepad_template*.png"):
        import glob
        self.template_paths = glob.glob(os.path.join(os.getcwd(), template_pattern))
        self.last_screenshot = None
        self.templates = []
        
        # Load all templates found
        if self.template_paths:
            for path in self.template_paths:
                tmpl = cv2.imread(path, cv2.IMREAD_COLOR)
                if tmpl is not None:
                    self.templates.append((os.path.basename(path), tmpl))
                    print(f"Loaded template: {os.path.basename(path)}")
                else:
                    print(f"Warning: Failed to load template from {path}")
        else:
            print(f"Warning: No templates found matching '{template_pattern}'")
            print("Please snip an image of the Notepad icon and save it as 'notepad_template.png' (or _small/_large) in this folder.")

    def capture_screenshot(self, save_path=None):
        """Captures the full desktop screenshot."""
        screenshot = pyautogui.screenshot()
        screenshot_np = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        self.last_screenshot = screenshot_np
        if save_path:
            cv2.imwrite(save_path, screenshot_np)
        return screenshot_np

    def find_icon(self, retry_attempts=3, delay=2):
        """
        Locates the target icon using Multiple Template Matching.
        Returns (x, y) center coordinates or None if not found.
        """
        if not self.templates:
            # Try reloading
            self.__init__()
            if not self.templates:
                print("Error: No reference images found.")
                return None

        # Minimize all windows to show desktop
        print("Minimizing windows to see desktop...")
        pyautogui.hotkey('win', 'd')
        time.sleep(2) 

        for attempt in range(retry_attempts):
            print(f"Grounding attempt {attempt + 1}/{retry_attempts} using {len(self.templates)} templates...")
            screenshot = self.capture_screenshot()
            
            best_overall_val = 0
            best_overall_loc = None
            best_template_name = ""
            best_dims = (0, 0)

            # Check all templates
            for name, tmpl in self.templates:
                try:
                    h, w = tmpl.shape[:2]
                    result = cv2.matchTemplate(screenshot, tmpl, cv2.TM_CCOEFF_NORMED)
                    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
                    
                    if max_val > best_overall_val:
                        best_overall_val = max_val
                        best_overall_loc = max_loc
                        best_template_name = name
                        best_dims = (w, h)
                except Exception as e:
                    print(f"Error matching template {name}: {e}")

            print(f"Best match: {best_template_name} with confidence {best_overall_val:.2f}")

            # Threshold check (0.9 preferred, but 0.8 is acceptable fallback)
            if best_overall_val >= 0.8:
                top_left = best_overall_loc
                w, h = best_dims
                center_x = top_left[0] + w // 2
                center_y = top_left[1] + h // 2
                
                print(f"Icon found using '{best_template_name}' at ({center_x}, {center_y})")
                return (center_x, center_y)
            else:
                print("Confidence too low. Icon not found.")
            
            if attempt < retry_attempts - 1:
                time.sleep(delay)
        
        print("Failed to find icon matching any template.")
        return None

    def annotate_detection(self, coords, output_path):
        """Save a screenshot with the detection marked."""
        if self.last_screenshot is None:
            self.capture_screenshot()
            
        img = self.last_screenshot.copy()
        if coords:
            cv2.circle(img, coords, 20, (0, 0, 255), 2) # Red circle
            cv2.line(img, (coords[0]-10, coords[1]), (coords[0]+10, coords[1]), (0,0,255), 2) # Crosshair
            cv2.line(img, (coords[0], coords[1]-10), (coords[0], coords[1]+10), (0,0,255), 2)
        
        cv2.imwrite(output_path, img)
        return output_path
