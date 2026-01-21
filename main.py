from grounding import VisualGrounding
from notepad_bot import NotepadBot
import time
import os

def main():
    print("--- Vision-Based Desktop Automation Starting ---")
    print("Resolution requirement: 1920x1080 (Targeting labels)")
    
    # Initialize components
    # Initialize components
    grounder = VisualGrounding(template_pattern="notepad_template*.png")
    bot = NotepadBot()
    
    # Fetch data
    print("Fetching posts from API...")
    posts = bot.fetch_posts(count=10)
    if not posts:
        print("No posts found or API error. Exiting.")
        return

    # Create deliverable directory
    deliverables_dir = "deliverables"
    if not os.path.exists(deliverables_dir):
        os.makedirs(deliverables_dir)

    # Main automation loop
    for i, post in enumerate(posts):
        print(f"\nProcessing post {i+1}/10 (ID: {post['id']})...")
        
        # 1. Capture screenshot and ground icon
        # We take a fresh screenshot for every cycle as required
        coords = grounder.find_icon(retry_attempts=3, delay=2)
        
        if not coords:
            print(f"Warning: Notepad icon not found. Will attempt fallback launch.")
            
        # 2. Capture annotated screenshot for first few cycles or specific needs
        if i < 3:
            anno_name = f"detection_step_{i+1}.png"
            grounder.annotate_detection(coords, os.path.join(deliverables_dir, anno_name))
            print(f"Annotated screenshot saved: {anno_name}")

        # 3. Launch and interact
        success = bot.run_cycle(post, coords)
        
        if success:
            print(f"Successfully processed post {post['id']}")
        else:
            print(f"Failed to process post {post['id']}")
            
        # Briefly wait before next cycle
        time.sleep(2)

    print("\n--- Automation Task Completed ---")

if __name__ == "__main__":
    main()
