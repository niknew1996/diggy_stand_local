import paramiko
import openpyxl

# ประกาศ Functions Telnet ผ่าน SSH
def test_telnet_via_ssh(ssh_client, dest_ip, port, timeout=5):
    try:
        stdin, stdout, stderr = ssh_client.exec_command(f"telnet {dest_ip} {port}", timeout=timeout)
        stdout.channel.settimeout(timeout)
        if "Connected" in stdout.read().decode():
            return "Success"
        else:
            return "Failed"
    except Exception as e:
        print(f"An error occurred during telnet: {e}")
        return "Failed"

# ประกาศ Functions main
def main(file_name):
    wb = openpyxl.load_workbook(file_name) # Loadfile Excel
    ws = wb.active 

    # รับค่า ssh_user และ ssh_password จากผู้ใช้
    ssh_user = input("Enter SSH Username: ")
    ssh_password = input("Enter SSH Password: ")
    ssh_port = 22

    for row in ws.iter_rows(min_row=3):  # Skip Header Row
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

        all_success = True
        for source_ip in range(int(source_start.split('.')[-1]), int(source_end.split('.')[-1]) + 1):
            source_ip_full = source_start.rsplit('.', 1)[0] + '.' + str(source_ip)

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

            # แกะค่า Destinations ip start กับ end ออกมา
            for dest_ip in range(int(dest_start.split('.')[-1]), int(dest_end.split('.')[-1]) + 1):
                dest_ip_full = dest_start.rsplit('.', 1)[0] + '.' + str(dest_ip)
                print(f"Testing telnet from {source_ip_full} to {dest_ip_full}:{port}")
                result = test_telnet_via_ssh(ssh_client, dest_ip_full, port) #เริ่ม telnet ผ่าน SSH
                print(f"Testing telnet result: {result}")
                if result == "Failed":
                    all_success = False #ถ้ามี failed 1 ตัว ก็ false หมดเลย

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
