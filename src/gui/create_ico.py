from PIL import Image
import os

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.join(os.path.abspath("."), "src", "gui")
    return os.path.join(base_path, relative_path)

img = Image.open(resource_path('.\logo-aplicativo.png'))
img.save('icone.ico', format='ICO')