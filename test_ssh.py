import paramiko
import openpyxl
import ipaddress
import getpass

# Function to test telnet connection via SSH
def test_telnet_through_ssh(ssh_client, dest_ip, port, timeout=5):
    try:
        command = f"echo quit | telnet {dest_ip} {port}"
        stdin, stdout, stderr = ssh_client.exec_command(command, timeout=timeout)
        output = stdout.read().decode()
        error = stderr.read().decode()
        if "Connected" in output or "Escape character is" in output:
            result = f"Telnet to {dest_ip}:{port} successful"
        else:
            result = f"Telnet to {dest_ip}:{port} failed: {error or output}"
        print(result)  # Print the result
        return result
    except Exception as e:
        result = f"Telnet to {dest_ip}:{port} failed: {e}"
        print(result)  # Print the result
        return result

# Function to validate IP addresses
def is_valid_ip(ip):
    try:
        ipaddress.ip_address(ip)
        return True
    except ValueError:
        return False

# Function to validate IP address range
def is_valid_ip_range(start_ip, end_ip):
    try:
        start_octet = int(start_ip.split('.')[-1])
        end_octet = int(end_ip.split('.')[-1])
        return 0 <= start_octet <= 255 and 0 <= end_octet <= 255 and is_valid_ip(start_ip) and is_valid_ip(end_ip)
    except (ValueError, AttributeError):
        return False

# Main function
def main(file_name):
    # Load Excel file
    wb = openpyxl.load_workbook(file_name)
    ws = wb.active 

    # Get SSH credentials from user
    ssh_user = input("Enter SSH Username: ")
    ssh_password = getpass.getpass("Enter SSH Password: ")  # Hide password input
    ssh_port = input("Enter SSH Port: ")

    # Iterate through rows starting from the third row
    for row in ws.iter_rows(min_row=3):
        if all(cell.value is None for cell in row):
            print("Row is Null, stopping process.")
            break

        # Extract IP ranges and port
        source_start, source_end, dest_start, dest_end = row[1].value, row[2].value, row[4].value, row[5].value

        # Validate IP address ranges
        if not (is_valid_ip_range(source_start, source_end) and is_valid_ip_range(dest_start, dest_end)):
            row[10].value = "Please provide valid IP addresses"
            continue  # Skip to next row if IP addresses are invalid

        try:  # Use port 80 if port is null
            port = int(row[7].value) if row[7].value else 80
        except ValueError:
            row[10].value = "Invalid Port"
            continue  # Skip to next row if port is invalid

        all_success = True
        source_network = source_start.rsplit('.', 1)[0]
        for source_ip in range(int(source_start.split('.')[-1]), int(source_end.split('.')[-1]) + 1):
            source_ip_full = f"{source_network}.{source_ip}"

            ssh_client = paramiko.SSHClient()
            ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            try:
                ssh_client.connect(source_ip_full, port=ssh_port, username=ssh_user, password=ssh_password)
                print(f"SSH connection to {source_ip_full} successful")
            except Exception as e:
                print(f"An error occurred when connecting to {source_ip_full}: {e}")
                row[10].value = "SSH Failed"
                all_success = False
                break  # Skip to next row if SSH connection fails

            # Iterate through destination IPs
            dest_network = dest_start.rsplit('.', 1)[0]
            for dest_ip in range(int(dest_start.split('.')[-1]), int(dest_end.split('.')[-1]) + 1):
                dest_ip_full = f"{dest_network}.{dest_ip}"
                print(f"Testing telnet from {source_ip_full} to {dest_ip_full}:{port}")
                result = test_telnet_through_ssh(ssh_client, dest_ip_full, port)  # Start telnet via SSH
                print(f"Testing telnet result: {result}")
                if "failed" in result:
                    all_success = False  # If any test fails, mark as false

            ssh_client.close()

        row[10].value = "Success" if all_success else "Failed"

    # Save the workbook with results
    try:
        wb.save(file_name)
    except PermissionError as e:
        print(f"Permission denied error: {e}. Please close any other application that might be using the file.")

if __name__ == "__main__":
    file_name = input("Filename of Excel (excel.xlsx): ")
    main(file_name)
