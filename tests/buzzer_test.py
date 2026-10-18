from gpiozero import OutputDevice, PWMOutputDevice
from time import sleep

print("Test 1: plain ON for 1 second")
pin = OutputDevice(18)
pin.on(); sleep(1); pin.off()
pin.close()
sleep(1)

print("Test 2: 2 kHz tone for 1 second")
pwm = PWMOutputDevice(18, frequency=2000)
pwm.value = 0.5; sleep(1); pwm.off()
