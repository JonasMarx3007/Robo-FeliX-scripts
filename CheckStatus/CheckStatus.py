import os
import sys
import time
import pywhatkit

def send_whatsapp_done_message(phone_number):
    message = "done"
    print(f"Sending WhatsApp message '{message}' to {phone_number}...")
    pywhatkit.sendwhatmsg_instantly(phone_number, message, wait_time=20, tab_close=True, close_time=5)

def get_exe_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def main():
    script_dir = get_exe_dir()
    done_file = os.path.join(script_dir, "done.txt")

    print(f"Executable directory is: {script_dir}")
    print(f"Looking for 'done.txt' at: {done_file}")

    while True:
        if os.path.isfile(done_file):
            send_whatsapp_done_message("+your_number")
            print("Message sent, exiting.")
            break
        else:
            print("'done.txt' not found. Checking again in 10 seconds...")
            time.sleep(10)

if __name__ == "__main__":
    main()
