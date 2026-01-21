import requests
import pyautogui
import time
import os
import pygetwindow as gw
import pyperclip
import socket
import urllib.request
import json

import shutil

class NotepadBot:
    def __init__(self, output_dir="Desktop/tjm-project"):
        # Disable FailSafe to allow interaction with icons on screen edges
        pyautogui.FAILSAFE = False
        
        # Resolve Desktop path
        self.output_dir = os.path.join(os.path.expanduser("~"), output_dir)
        
        # Cleanup existing directory
        if os.path.exists(self.output_dir):
            print(f"Cleaning up existing directory: {self.output_dir}")
            shutil.rmtree(self.output_dir)
            time.sleep(1) # Wait for filesystem to catch up
            
        os.makedirs(self.output_dir)
            
        self.api_url = "https://jsonplaceholder.typicode.com/posts"
        
        # Create a session for better connection handling
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Connection': 'keep-alive',
            'Accept': 'application/json'
        })

    def test_connection(self):
        """Test basic network connectivity."""
        try:
            # Test DNS resolution
            socket.getaddrinfo('jsonplaceholder.typicode.com', 443)
            print("✓ DNS resolution successful")
            
            # Test basic connection
            sock = socket.create_connection(('jsonplaceholder.typicode.com', 443), timeout=10)
            sock.close()
            print("✓ TCP connection successful")
            return True
        except Exception as e:
            print(f"✗ Connection test failed: {e}")
            return False

    def fetch_posts(self, count=10):
        """Fetch posts directly from the API with enhanced error handling."""
        
        # Method 1: Fresh requests without session (avoid connection pooling issues)
        for attempt in range(3):
            try:
                print(f"Connecting to API (Attempt {attempt+1}/3)...")
                
                # Create fresh connection each time
                import urllib3
                urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
                
                response = requests.get(
                    self.api_url, 
                    timeout=20,
                    verify=False,  # Disable SSL verification
                    headers={
                        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                        'Accept': 'application/json',
                        'Connection': 'close'  # Don't try to reuse connection
                    }
                )
                response.raise_for_status()
                posts = response.json()
                print(f"Successfully fetched {len(posts)} posts.")
                return posts[:count]
            except requests.exceptions.ConnectionError as e:
                print(f"Connection error on attempt {attempt+1}: {e}")
                if "10054" in str(e):
                    print("-> Connection forcibly closed. Trying alternative method...")
            except Exception as e:
                print(f"Request failed on attempt {attempt+1}: {e}")
            
            if attempt < 2:
                time.sleep(3)

        # Method 2: Urllib with SSL context (Backup for ConnectionReset errors)
        try:
            print("Attempting backup connection via urllib...")
            import ssl
            
            # Create SSL context that's more permissive
            context = ssl._create_unverified_context()
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                'Connection': 'close'
            }
            req = urllib.request.Request(self.api_url, headers=headers)
            with urllib.request.urlopen(req, timeout=20, context=context) as url:
                data = json.loads(url.read().decode())
                print(f"Successfully fetched {len(data)} posts via Urllib.")
                return data[:count]
        except Exception as e:
            print(f"Urllib HTTPS failed: {e}")

        # Method 3: Try HTTP instead of HTTPS as last resort
        try:
            print("Trying HTTP (non-secure) connection...")
            http_url = "http://jsonplaceholder.typicode.com/posts"
            response = requests.get(
                http_url, 
                timeout=20,
                headers={'Connection': 'close'}
            )
            response.raise_for_status()
            posts = response.json()
            print(f"Successfully fetched {len(posts)} posts via HTTP.")
            return posts[:count]
        except Exception as e:
            print(f"HTTP fallback failed: {e}")
        
        # Method 4: Alternative API endpoint using requests-html or direct socket
        try:
            print("Final attempt: Using raw socket connection...")
            import socket
            import ssl
            
            # Build HTTP request manually
            host = "jsonplaceholder.typicode.com"
            request = f"GET /posts HTTP/1.1\r\nHost: {host}\r\nUser-Agent: Mozilla/5.0\r\nConnection: close\r\n\r\n"
            
            # Create socket and wrap with SSL
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(20)
            sock.connect((host, 443))
            
            context = ssl._create_unverified_context()
            ssock = context.wrap_socket(sock, server_hostname=host)
            ssock.send(request.encode())
            
            # Read response
            response_data = b""
            while True:
                chunk = ssock.recv(4096)
                if not chunk:
                    break
                response_data += chunk
            
            ssock.close()
            
            # Parse response
            response_text = response_data.decode('utf-8')
            json_start = response_text.find('[')
            if json_start != -1:
                json_data = response_text[json_start:]
                posts = json.loads(json_data)
                print(f"Successfully fetched {len(posts)} posts via raw socket.")
                return posts[:count]
        except Exception as e:
            print(f"Raw socket connection failed: {e}")

        print("CRITICAL: All API connection attempts failed. Using generated mock data.")
        return [
            {
                "id": i, 
                "title": f"Mock Post {i}", 
                "body": f"This is the body validation for Mock Post {i}."
            }
            for i in range(1, count + 1)
        ]

    def wait_for_notepad(self, timeout=15):
        """STRICTEST window validation to avoid IDEs."""
        start_time = time.time()
        print("Searching ONLY for the Windows Notepad application window...")
        
        while time.time() - start_time < timeout:
            all_windows = gw.getAllWindows()
            for win in all_windows:
                title = win.title
                
                # Notepad windows usually have "- Notepad" or are just "Notepad"
                # We strictly exclude any IDE or Script names
                is_notepad = "Notepad" in title
                is_false_positive = any(x in title for x in [" - Antigravity", "Visual Studio Code", "notepad_bot.py", ".py", ".toml"])
                
                if is_notepad and not is_false_positive:
                    try:
                        win.activate()
                        if win.isMinimized:
                            win.restore()
                        time.sleep(2.0) # Longer wait for focus
                        print(f"FOUND VALID NOTEPAD: '{title}'")
                        return True
                    except Exception:
                        continue
            time.sleep(1)
        return False

    def process_post(self, post):
        """Paste data for 100% accuracy and save to folder."""
        title = post.get('title')
        body = post.get('body')
        post_id = post.get('id')
        filename = f"post_{post_id}.txt"
        file_path = os.path.normpath(os.path.join(self.output_dir, filename))

        print(f"Processing Post {post_id}...")

        # 1. Clear window
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
        time.sleep(1.0)

        # 2. Paste Content (Format as requested)
        full_text = f"Title: {title}\n\n{body}"
        pyperclip.copy(full_text)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1.5)

        # 3. Save As Menu
        print(f"Saving {filename}...")
        pyautogui.press('alt')
        time.sleep(0.5)
        pyautogui.press('f')
        time.sleep(0.5)
        pyautogui.press('a')
        time.sleep(2.5) 

        # 4. Input Path and Enter
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
        time.sleep(0.8)
        pyperclip.copy(file_path)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.8)
        pyautogui.press('enter')
        time.sleep(2.5)

        # 5. Confirmation dialog safety
        confirm = gw.getWindowsWithTitle('Confirm Save As')
        if confirm:
            print("Confirm Save As detected. Overwriting...")
            pyautogui.press('y')
            time.sleep(1.5)
            
        # 5b. Bypass unexpected pop-ups (User Request)
        # If any OTHER window is active (not Notepad, not Save As), press ESC to ignore/dismiss.
        try:
            active_win = gw.getActiveWindow()
            if active_win:
                t = active_win.title.lower()
                if "notepad" not in t and "save as" not in t and "confirm save as" not in t:
                    print(f"Unexpected blocking window '{active_win.title}'. Bypassing with ESC...")
                    pyautogui.press('esc')
                    time.sleep(1.0)
        except Exception as e:
            print(f"Popup bypass check failed: {e}")

        # 6. Exit
        # User requested Ctrl+W
        pyautogui.hotkey('ctrl', 'w')
        time.sleep(1.5)
        
        if os.path.exists(file_path):
            print(f"SUCCESS: Saved {file_path}")
            return True
        else:
            print(f"ERROR: File verification failed for {filename}")
            return False

    def run_cycle(self, post, grounding_coords):
        """Physical mouse interaction and activation."""
        # Launch Logic
        if grounding_coords:
            # Physical mouse move for visual feedback
            print(f"Launching Notepad via icon at {grounding_coords}...")
            pyautogui.moveTo(grounding_coords[0], grounding_coords[1], duration=0.8)
            pyautogui.doubleClick()
            # Move mouse slightly to prevent hover tooltip from obstructing next screenshot
            time.sleep(0.2)
            pyautogui.moveRel(200, 0) 
        else:
            # Fallback: Win+R
            print("Icon not found. Launching via Win+R...")
            pyautogui.hotkey('win', 'r')
            time.sleep(0.5)
            pyperclip.copy("notepad")
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            pyautogui.press('enter')
        
        result = False
        if self.wait_for_notepad():
            result = self.process_post(post)
        else:
            print("NOTEPAD NOT FOUND. SAFETY ABORT.")
            result = False
            
        # print("Resetting mouse position...")
        # pyautogui.moveTo(100, 100)
        return result