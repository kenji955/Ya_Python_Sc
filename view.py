import eel
import sys
import desktop
import ebay_rakuma

app_name = "web"
end_point = "index.html"
size = (500, 500)


@ eel.expose
def main():
    ebay_rakuma.main()
    sys.exit(0)


desktop.start(app_name, end_point, size)
sys.exit(0)
