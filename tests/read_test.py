import select, sys, termios, time, tty
import serial

PORT = sys.argv[1] if len(sys.argv) > 1 else "/dev/serial0"

ser = serial.Serial(PORT, 9600, timeout=0)    # non-blocking: select() does the waiting
fd = sys.stdin.fileno()
old = termios.tcgetattr(fd)
tty.setcbreak(fd)                              # read single keys without Enter

print("Waiting for card... (press 0 to quit)")
buf, last_id, last_time = b"", None, 0

try:
    while True:
        # Sleep until the reader sends data or a key is pressed
        ready, _, _ = select.select([ser, sys.stdin], [], [])

        if sys.stdin in ready:
            ch = sys.stdin.read(1)
            if ch in ("0", ""):                # "" means the terminal closed
                break

        if ser not in ready:
            continue
        for i in ser.read(ser.in_waiting or 1):
            b = bytes([i])
            if b == b"\x02":                   # start of frame
                buf = b""
            elif b == b"\x03":                 # end of frame
                frame, buf = buf.decode("ascii", "ignore"), b""
                try:
                    data, chk = frame[:10], int(frame[10:12], 16)
                    x = 0
                    for j in range(0, 10, 2):
                        x ^= int(data[j:j+2], 16)
                    if len(frame) != 12 or x != chk:
                        print("Bad read:", frame); continue
                except ValueError:
                    print("Bad read:", frame); continue
                n = int(data[2:], 16)          # card ID (last 8 hex chars)
                now = time.time()
                if n == last_id and now - last_time < 2:   # module repeats while card is held
                    last_time = now; continue
                last_id, last_time = n, now
                print(f"Raw: {data}  ID: {n:010d}  Wiegand: {(n >> 16) & 0xFF:03d},{n & 0xFFFF:05d}")
            else:
                buf += b
except KeyboardInterrupt:
    pass
finally:
    termios.tcsetattr(fd, termios.TCSADRAIN, old)   # restore normal typing
    ser.close()
    print("Stopped.")