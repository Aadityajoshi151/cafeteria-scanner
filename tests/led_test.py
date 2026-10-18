import select, sys, termios, tty
from gpiozero import LED

red, green = LED(19), LED(26)          # pins 35 and 37; common cathode to GND
steps = [("RED", 1, 0), ("GREEN", 0, 1), ("OFF", 0, 0)]

fd = sys.stdin.fileno()
old = termios.tcgetattr(fd)
tty.setcbreak(fd)                      # read single keys without Enter
print("Cycling colours every 2 s... (press 0 to quit)")
try:
    i = 0
    while True:
        name, r, g = steps[i % len(steps)]
        red.value, green.value = r, g
        print(name)
        i += 1
        if select.select([sys.stdin], [], [], 2)[0] and sys.stdin.read(1) in ("0", ""):
            break
except KeyboardInterrupt:
    pass
finally:
    termios.tcsetattr(fd, termios.TCSADRAIN, old)   # restore normal typing
    red.close(); green.close()
    print("Stopped.")