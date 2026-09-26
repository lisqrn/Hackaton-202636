import os
import subprocess
import re

EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
BASE_DIR = r"C:\Users\djahe\OneDrive\Documents\Alisa\skynet\Hackaton-202636"
HTML_PATH = os.path.join(BASE_DIR, "presentation.html")
OUTPUT_DIR = os.path.join(BASE_DIR, "slide_images")
os.makedirs(OUTPUT_DIR, exist_ok=True)

with open(HTML_PATH, "r", encoding="utf-8") as f:
    html_content = f.read()

# Strip any script timer
clean_html = re.sub(r"<script>.*?</script>", "", html_content, flags=re.DOTALL)

for i in range(1, 9):
    export_css = f"""
    <style>
      html, body {{
        width: 1920px !important;
        height: 1080px !important;
        overflow: hidden !important;
      }}
      .deck-viewport {{
        width: 1920px !important;
        height: 1080px !important;
        max-width: 1920px !important;
        max-height: 1080px !important;
        padding: 45px 75px 45px !important;
      }}
      .bottom-control-bar, .notes-drawer, .deck-progress-bar {{
        display: none !important;
      }}
      .slide {{
        display: none !important;
        opacity: 0 !important;
      }}
      .slide[data-slide="{i}"] {{
        display: flex !important;
        opacity: 1 !important;
        transform: none !important;
      }}
    </style>
    """
    slide_html = clean_html.replace("</head>", f"{export_css}</head>")
    slide_html = slide_html.replace('id="slideIndexLabel">01 / 08<', f'id="slideIndexLabel">0{i} / 08<')
    
    temp_file = os.path.join(OUTPUT_DIR, f"temp_slide_{i}.html")
    with open(temp_file, "w", encoding="utf-8") as f:
        f.write(slide_html)
        
    png_file = os.path.join(OUTPUT_DIR, f"slide_{i}.png")
    cmd = [
        EDGE_PATH,
        "--headless",
        "--no-sandbox",
        "--disable-gpu",
        f"--screenshot={png_file}",
        "--window-size=1920,1080",
        temp_file
    ]
    print(f"Rendering Slide {i}...")
    subprocess.run(cmd, capture_output=True)
    if os.path.exists(png_file):
        print(f"Slide {i} rendered: {os.path.getsize(png_file)} bytes")
    else:
        print(f"Failed to render slide {i}")

print("All slides processed successfully!")
