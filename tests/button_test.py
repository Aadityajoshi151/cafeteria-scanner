import sys, termios, tty
from gpiozero import Button

buttons = {
    "VEG": Button(17, bounce_time=0.05),
    "NON-VEG": Button(27, bounce_time=0.05),
    "COMPLETE": Button(22, bounce_time=0.05),
}

for name, btn in buttons.items():
    btn.when_pressed = lambda n=name: print(f"{n} pressed")
    btn.when_released = lambda n=name: print(f"{n} released")

fd = sys.stdin.fileno()
old = termios.tcgetattr(fd)
tty.setcbreak(fd)                      # read single keys without Enter

print("Press buttons... (press 0 to quit)")
try:
    while True:
        ch = sys.stdin.read(1)
        if ch in ("0", ""):            # "" means the terminal closed
            break
except KeyboardInterrupt:
    pass
finally:
    termios.tcsetattr(fd, termios.TCSADRAIN, old)   # restore normal typing
    for b in buttons.values():
        b.close()
    print("Stopped.")