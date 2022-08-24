import keyboard
import threading
import time

def KeyboardBlock():
    print("Blocking Keyboard")
    for i in range(150):
        keyboard.block_key(i)
    time.sleep(10)
    print("Unblocking keyboard")

#KBThread=threading.Thread(target=KeyboardBlock(),name="KeyboardBlock")
#KBThread.start()
