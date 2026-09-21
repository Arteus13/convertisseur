import sys
import os
import json
import struct
import subprocess

# Fix streams
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")

def read_message():
    # Chrome envoie 4 octets pour la longueur du message JSON
    raw_length = sys.stdin.buffer.read(4)
    if len(raw_length) == 0:
        sys.exit(0)
    msg_length = struct.unpack('@I', raw_length)[0]
    data = sys.stdin.buffer.read(msg_length).decode('utf-8')
    return json.loads(data)

def send_message(msg_dict):
    msg_json = json.dumps(msg_dict).encode('utf-8')
    sys.stdout.buffer.write(struct.pack('@I', len(msg_json)))
    sys.stdout.buffer.write(msg_json)
    sys.stdout.buffer.flush()

if __name__ == "__main__":
    try:
        msg = read_message()
        file_path = msg.get("filePath")
        if file_path and os.path.exists(file_path):
            local_app_data = os.environ.get("LOCALAPPDATA", "")
            exe_path = os.path.join(local_app_data, "Programs", "ConvertisseurUniversel", "convertisseur.exe")
            
            if os.path.exists(exe_path):
                subprocess.Popen([exe_path, file_path])
                send_message({"status": "success", "file": file_path})
            else:
                send_message({"status": "error", "message": "convertisseur.exe non installe"})
    except Exception as e:
        pass
