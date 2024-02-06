from tqdm import tqdm
import subprocess
import re
import threading
import openpyxl
import platform   

def read_excel_column(file_path, sheet_name, column):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook[sheet_name]
    data_column = [cell.value for cell in sheet[column][1:]]
    workbook.close()
    return data_column

def ping_host(hostname):
    param = '-n' if platform.system().lower()=='windows' else '-c'
    str_time_out = int(config_list[4]) * 1000 
    time_out = str(str_time_out)
    ping_times = str(config_list[5])

    command = ["ping", param, ping_times,"-w", time_out, hostname]
    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    output, error = process.communicate()
    if process.returncode == 0:
        delay_match = re.search(r"(\d+\.?\d*)ms", output.decode("utf-8", errors='ignore'))
        if delay_match:
            delay_ms = float(delay_match.group(1))
            if config_list[2] == 'yes' :
                print(f"{hostname} is up! Delay: {delay_ms} ms")
        else:
            print(f"Unable to extract delay information for {hostname}")
    else:
        if config_list[3] == 'yes' :
            print(f"{hostname} is down!")

    if config_list[6] == 'yes' :
        progress_bar.update(1)

if __name__ == "__main__":
    config_list = read_excel_column('config.xlsx', 'config', 'B')  ## 這行不能改，用excel 做config 檔

    for runs in range(int(config_list[7])) :
        hostname_list = read_excel_column('config.xlsx', config_list[0],config_list[1])

        total_delay = 0
        threads = []
        if config_list[6] == 'yes' :
            progress_bar = tqdm(total=len(hostname_list), desc="Pinging hosts", unit="host")
        
        print("---------------Processing(Please Wait)---------------")

        for hostname in hostname_list:
            thread = threading.Thread(target=ping_host, args=(hostname,)) 
            threads.append(thread)
            thread.start()

        for thread in threads:
            thread.join()

        printrun = int(runs) + 1 
        print(f"        ---------------End {printrun}---------------        ")