import asyncio
import paramiko
import openpyxl
from telnetlib3 import open_connection

# ประกาศ Functions Telnet
async def test_telnet(telnet_host, telnet_port, timeout=5):
    try:
        reader, writer = await asyncio.wait_for(open_connection(telnet_host, port=telnet_port), timeout=timeout)
        await asyncio.wait_for(reader.read(1), timeout=timeout)
        writer.close()
        return "Success"
    except (asyncio.TimeoutError, Exception):
        return "Failed"

# ประกาศ Functions main
def main(file_name):
    wb = openpyxl.load_workbook(file_name) # Loadfile Excel
    ws = wb.active 

    for row in ws.iter_rows(min_row=2):  # Skip Header Row
        if all(cell.value is None for cell in row):
            print("Row is Null stop it")
            break

        # ตรวจสอบค่า IP range และพอร์ต
        source_start, source_end, dest_start, dest_end = row[1].value, row[2].value, row[4].value, row[5].value
        try: # is port null use 80
            port = int(row[7].value) if row[7].value else 80
        except ValueError:
            row[10].value = "Invalid Port"
            continue

        ssh_user, ssh_password = row[12].value, row[13].value
        ssh_port = 22  # พอร์ต SSH ปกติ


        all_success = True #ถ้ามี null ในตัวแปรจะ false ทันที
        #แกะค่า source ip start กับ end ออกมา
        for source_ip in range(int(source_start.split('.')[-1]), int(source_end.split('.')[-1]) + 1):
            source_ip_full = source_start.rsplit('.', 1)[0] + '.' + str(source_ip)
            #remote sshclient เข้า source ip start

            ssh_client = paramiko.SSHClient()
            ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            try:
                ssh_client.connect(source_ip_full, port=ssh_port, username=ssh_user, password=ssh_password)
                print(f"SSH connection to {source_ip_full} successful")
            except Exception as e:
                print(f"An error occurred when connecting to {source_ip_full}: {e}")
                row[10].value = "SSH Failed"
                all_success = False
                break
                #แกะค่า Destinations ip start กับ end ออกมา
            for dest_ip in range(int(dest_start.split('.')[-1]), int(dest_end.split('.')[-1]) + 1):
                dest_ip_full = dest_start.rsplit('.', 1)[0] + '.' + str(dest_ip)
                print(f"Testing telnet from {source_ip_full} to {dest_ip_full}:{port}")
                result = asyncio.run(test_telnet(dest_ip_full, port)) #เริ่ม telnet
                print(f"Testing telnet result: {result}")
                if result == "Failed":
                    all_success = False #ถ้ามี failed 1 ตัว ก็ false หมดเลย
              #  row[10].value = result #บันทึกลง Colum K

            ssh_client.close()

        row[10].value = "Success" if all_success else "Failed"

    # บันทึกผลลัพธ์ลงไฟล์ Excel
    try:
        wb.save(file_name)
    except PermissionError as e:
        print(f"Permission denied error: {e}. Please close any other application that might be using the file.")

if __name__ == "__main__":
    file_name = input("Filename of Excel (excel.xlsx): ")
    main(file_name)
