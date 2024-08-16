import os
import threading
import itertools
import time

message = ""
message_lock = threading.Lock()

def clear_console():
    os.system('cls' if os.name == 'nt' else 'clear')

def task():
    global message
    for i in range(10):
        with message_lock:
            message = f"Chargement de la tâche {i+1}/10"
        time.sleep(0.5)

def spinner(stop_event):
    global message
    for frame in itertools.cycle('|/-\\'):
        if stop_event.is_set():
            break
        clear_console()
        with message_lock:
            print(frame + " " + message)
        time.sleep(0.1)

stop_event = threading.Event()
spinner_thread = threading.Thread(target=spinner, args=(stop_event,))

try:
    spinner_thread.start()
    task()
finally:
    stop_event.set()
    spinner_thread.join()
    clear_console()
    print("Chargement terminé!") 

