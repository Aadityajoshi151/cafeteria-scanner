from luma.core.interface.serial import spi
from luma.core.render import canvas
from luma.oled.device import ssd1306

serial = spi(port=0, device=0, gpio_DC=24, gpio_RST=25)
device = ssd1306(serial)

with canvas(device) as draw:
    draw.rectangle(device.bounding_box, outline="white")
    draw.text((8, 12), "Cafeteria scanner", fill="white")
    draw.text((8, 36), "OLED test OK", fill="white")

input("Showing text. Press Enter to exit...")
