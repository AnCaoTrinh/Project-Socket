import socket
import base64
import os
import random
import string
from email import policy
from email.parser import BytesParser
import re
def get_content_type(file_path):
    # """Xác định Content-Type cho một định dạng file cụ thể."""
    file_extension = os.path.splitext(file_path)[1].lower()
    # print(file_extension)
    if file_extension == '.pdf':
        return 'application/pdf'
    elif file_extension == '.docx':
        return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    elif file_extension == '.jpg':
        return 'image/jpeg'
    elif file_extension == '.zip':
        return 'application/zip'
    else:
        return 'application/octet-stream'  # Nếu không xác định được, sử dụng octet-stream

def generate_boundary():
    return ''.join(random.choice(string.ascii_letters + string.digits) for _ in range(30))

def get_file_size(file_path):
    return os.path.getsize(file_path)

def send_mail(smtp_server, smtp_port, from_address, to_addresses, cc_addresses=None, bcc_addresses=None, subject='', body='', attachments=None):
    try:
        # Kết nối đến SMTP server
        client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        client_socket.connect((smtp_server, smtp_port))
        recv = client_socket.recv(1024).decode()
        # print(recv)

        # Gửi lệnh EHLO để bắt đầu phiên làm việc
        client_socket.send(b'EHLO [127.0.0.1]\r\n')
        recv1 = client_socket.recv(1024).decode()
        # print(recv1)

        # Gửi thông tin email
        client_socket.send(f'MAIL FROM: <{from_address}>\r\n'.encode())
        recv_mail_from = client_socket.recv(1024).decode()
        # print(recv_mail_from)

        # Gửi thông tin người nhận
        recipients = to_addresses + (cc_addresses or []) + (bcc_addresses or [])
        for recipient in recipients:
            client_socket.send(f'RCPT TO: <{recipient}>\r\n'.encode())
            recv_rcpt_to = client_socket.recv(1024).decode()
            # print(recv_rcpt_to)

        # Gửi lệnh DATA
        client_socket.send(b'DATA\r\n')
        recv_data = client_socket.recv(1024).decode()
        # print(recv_data)

        # Tạo boundary string
        boundary_string = generate_boundary()

        # Chuẩn bị và gửi nội dung email
        client_socket.send(f'Subject: {subject}\r\n'.encode())
        client_socket.send(f'From: {from_address}\r\n'.encode())
        client_socket.send(f'To: {", ".join(to_addresses)}\r\n'.encode())
        if cc_addresses:
            client_socket.send(f'Cc: {", ".join(cc_addresses)}\r\n'.encode())
        client_socket.send(f'Content-Type: multipart/mixed; boundary={boundary_string}\r\n'.encode())
        client_socket.send('\r\n'.encode())  # Kết thúc phần header, bắt đầu nội dung email

        # Gửi body email
        client_socket.send(f'--{boundary_string}\r\n'.encode())
        client_socket.send(f'Content-Type: text/plain\r\n\r\n{body}\r\n'.encode())

       # Đính kèm file nếu có
        if attachments:
            for attachment_path in attachments:
                filename = os.path.basename(attachment_path)
                file_size = get_file_size(attachment_path)

                if file_size > 3 * 1024 * 1024:  # Kiểm tra giới hạn dung lượng file là <= 3MB
                    print(f"File {filename} exceeds the size limit of 3MB. Skipped.")
                    continue

                # Xác định Content-Type cho định dạng file
                content_type = get_content_type(attachment_path)

                client_socket.send(f'--{boundary_string}\r\n'.encode())
                client_socket.send(f'Content-Type: {content_type}; name="{filename}"\r\n'.encode())
                client_socket.send(f'Content-Disposition: attachment; filename="{filename}"\r\n'.encode())

                with open(attachment_path, 'rb') as file:
                    attachment_content = base64.b64encode(file.read()).decode()
                    client_socket.send(f'Content-Transfer-Encoding: base64\r\n\r\n'.encode())
                    for i in range(0, len(attachment_content), 72):
                        line = attachment_content[i:i+72]
                        client_socket.send(f'{line}\r\n'.encode())
                        
        # Kết thúc nội dung email
        client_socket.send(f'--{boundary_string}--\r\n'.encode())
        client_socket.send('.\r\n'.encode())  # Dấu chấm kết thúc quá trình truyền dữ liệu thư
        recv_data = client_socket.recv(1024).decode()
        # print(recv_data)

    except Exception as e:
        print(f"An error occurred while sending email: {e}")

    finally:
        try:
            client_socket.send(b'QUIT\r\n')
        except:
            pass
        client_socket.close()


def files_in_folder(file_name, folder_path):
    # Kiểm tra xem đường dẫn đến thư mục có tồn tại không
    if not os.path.exists(folder_path):
        # print(f"Thư mục '{folder_path}' không tồn tại.")
        return False
    # Lặp qua tất cả các file trong thư mục
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)

        # Kiểm tra xem đường dẫn là file hay thư mục
        if os.path.isfile(file_path):
            # Xử lý file ở đây, ví dụ: in ra tên file
            if file_name == filename:
                return True
        elif os.path.isdir(file_path):
            # Nếu là thư mục, bạn có thể gọi đệ quy để xử lý thư mục con
            # print(f"Thư mục: {filename}")
            if files_in_folder(file_name, file_path):
                return True
    return False

def read_msg_content(msg_content):
    try:
        # Giải mã base64 và phân tích cú pháp
        msg = BytesParser(policy=policy.default).parsebytes(msg_content)
        from_mail2 = str(msg['From'])
        subject2 = str(msg['Subject'])
        body2 = str(msg.get_body().get_content())
        return from_mail2, subject2, body2
    except Exception as e:
        print(f"An error occurred: {e}")

def get_mail(pop3_server, pop3_port, username, password, folder_path, config, filter):
    try:
        # Kết nối đến máy chủ POP3
        client_socket = socket.create_connection((pop3_server, pop3_port))

        # Nhận và in thông báo kết nối từ máy chủ
        recv_data = client_socket.recv(1024).decode()
        # print(recv_data)

        # Gửi lệnh USER và PASS để xác thực
        client_socket.send(f'USER {username}\r\n'.encode())
        recv_user = client_socket.recv(1024).decode()
        # print(recv_user)

        client_socket.send(f'PASS {password}\r\n'.encode())
        recv_pass = client_socket.recv(1024).decode()
        # print(recv_pass)

        # Gửi lệnh STAT để lấy thông tin trạng thái hộp thư
        client_socket.send(b'STAT\r\n')
        recv_stat = client_socket.recv(1024).decode()
        # print(recv_stat)

        # Gửi lệnh LIST để lấy danh sách các email và kích thước của chúng
        client_socket.send(b'LIST\r\n')
        recv_list = client_socket.recv(1024).decode()
        # print(recv_list)
        # Gửi lệnh UIDL để lấy msg 
        client_socket.send(b'UIDL\r\n')
        recv_msg = client_socket.recv(1024).decode()
        # bo? qua dong +OK 
        for i in range(5, len(recv_msg), 25):
            if not files_in_folder(recv_msg[i + 2: i + 2 + 21], folder_path):
                # Gửi lệnh RETR để lấy nội dung của email
                if recv_msg[i] == '.' :
                    break
                client_socket.send(f'RETR {recv_msg[i]}\r\n'.encode())
                email_content = b""
                while True:
                    recv_rtr = client_socket.recv(1024)
                    email_content += recv_rtr
                    if recv_rtr.endswith(b'\r\n.\r\n'):
                        # read file msg neu 
                        from_mail2, subject2, body2 = read_msg_content(email_content)
                        for x in filter["From"] :
                            if from_mail2 == x :
                                email_file_path = os.path.join(folder_path[config[x]], recv_msg[i + 2: i + 2 + 21])
                                with open(email_file_path, 'wb') as email_file:
                                    email_file.write(email_content)
                                break
                        for x in filter["Subject"] :
                            if subject2.find(x[1: len(x) - 1]) != -1 :
                                email_file_path = os.path.join(folder_path[config[x[1: len(x) - 1]]], recv_msg[i + 2: i + 2 + 21])
                                with open(email_file_path, 'wb') as email_file:
                                    email_file.write(email_content)
                                break
                        for x in filter["Content"] :
                            if body2.find(x[1: len(x) - 1]) != -1 :
                                email_file_path = os.path.join(folder_path[config[x[1: len(x) - 1]]], recv_msg[i + 2: i + 2 + 21])
                                with open(email_file_path, 'wb') as email_file:
                                    email_file.write(email_content)
                                break
                        break

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        # Đóng kết nối
        try:
            client_socket.send(b'QUIT\r\n')
        except:
            pass
        client_socket.close()



def read_msg_file(msg_file_path):
    try:
        # Đọc nội dung của file .msg
        with open(msg_file_path, 'rb') as file:
            msg_content = file.read()

        # Giải mã base64 và phân tích cú pháp
        msg = BytesParser(policy=policy.default).parsebytes(msg_content)
        from_mail2 = str(msg['From'])
        subject2 = str(msg['Subject'])
        # In thông tin email
        print(f"Subject: {msg['Subject']}")
        print(f"From: {msg['From']}")
        print(f"To: {msg['To']}")
        print(f"Cc: {msg['Cc']}")
        print(f"Bcc: {msg['Bcc']}")

        # In nội dung email
        print("\nBody:")
        body2 = str(msg.get_body().get_content())
        print(body2)

        # Tạo thư mục để lưu đính kèm (nếu chưa tồn tại)
        os.makedirs(output_folder, exist_ok=True)

        # Xử lý đính kèm nếu có
        for idx, part in enumerate(msg.iter_parts()):
            if part.get_content_disposition() == 'attachment':
                number = int(input("Trong email này có attached file, bạn có muốn save không(1: có , other: không): "))
                if number == 1 :
                    output_folder = input("Cho biết đường dẫn bạn muốn lưu: ")
                    attachment_content = part.get_payload(decode=True)
                    # Sử dụng tên file từ 'filename' trong 'Content-Disposition'
                    disposition = part.get("Content-Disposition")
                    if disposition:
                        match = re.search(r'filename="(.+)"', disposition)
                        if match:
                            attachment_filename = match.group(1)
                        else:
                            attachment_filename = f"attachment_{idx + 1}"
                    else:
                        attachment_filename = f"attachment_{idx + 1}"
                    
                    # Đường dẫn đến file mới
                    attachment_file_path = os.path.join(output_folder, attachment_filename)

                    # Ghi nội dung của đính kèm vào file mới
                    with open(attachment_file_path, 'wb') as attachment_file:
                        attachment_file.write(attachment_content)

                    print(f"\nAttachment saved: {attachment_file_path}")
                else :
                    break
        return from_mail2, subject2, body2
    except Exception as e:
        print(f"An error occurred: {e}")

def menu() :
    print("""Vui lòng chọn Menu:
1. Để gửi email
2. Để xem danh sách các email đã nhận
3. Thoát""")

def mailbox() :
    print("""Đây là danh sách các folder trong mailbox của bạn:
1. Inbox
2. Project
3. Important
4. Work
5. Spam""")

if __name__ == '__main__' :
    # doc config
    config = {}
    with open("C:\\Users\\ASUS ZenBook\\OneDrive\\Máy tính\\python\\Project-Socket\\config.txt", 'r') as file :
        line = file.readline() # doc General:
        line = file.readline()
        while line:
            # loại bỏ những khoảng trắng 
            line = line.strip()

            if line[0] == '#' or not line :
                line = file.readline()
                continue
            if line == 'Filter:' :
                break
            key, value = line.split(": ")
            config[key] = value
            line = file.readline()
        filter = {}
        line = file.readline()
        while line :
            # loại bỏ những khoảng trắng 
            line = line.strip()

            if line[0] == '#' or not line :
                line = file.readline()
                continue

            first , value = line.split(" - To folder: ")

            buffer, Keys = first.split(": ")

            keys = Keys.split(", ")
            filter[buffer] = keys
            for x in keys :
                config[x] = value
            line = file.readline()
    # print(config)
    ls = config['Username'].split(" ")
    username = ' '.join(ls[0: len(ls) - 1])
    from_mail = ls[len(ls) - 1]
    folder_path = {}
    temp = {}
    temp['1'] = 'Inbox'
    temp['2'] = 'Project'
    temp['3'] = 'Important'
    temp['4'] = 'Work'
    temp['5'] = 'Spam'

    folder_path['Important'] = "C:\\Users\\ASUS ZenBook\\OneDrive\\Máy tính\\python\\Project-Socket\\Important"
    folder_path['Project'] = "C:\\Users\\ASUS ZenBook\\OneDrive\\Máy tính\\python\\Project-Socket\\Project"
    folder_path['Work'] = "C:\\Users\\ASUS ZenBook\\OneDrive\\Máy tính\\python\\Project-Socket\\Work"
    folder_path['Spam'] = "C:\\Users\\ASUS ZenBook\\OneDrive\\Máy tính\\python\\Project-Socket\\Spam"
    folder_path['Inbox'] = "C:\\Users\\ASUS ZenBook\\OneDrive\\Máy tính\\python\\Project-Socket\\Inbox"
    # print(from_mail)
    while True :
        menu()
        choice = input("Bạn chọn: ")
        if choice == '1' :
            print("Đây là thông tin soạn email: (nếu không điền vui lòng nhấn enter để bỏ qua)")
            to_email = input("To: ").split(", ")
            cc_email = input("CC: ").split(", ")
            bcc_email = input("BCC: ").split(", ")
            subject = input("Subject: ")
            content = input("Content: ")
            attachment_file = int(input("Có gửi kèm file (1. có, 2. không): "))
            if attachment_file == 1 :
                number_of_file = int(input("Số lượng file muốn gửi: "))
                attachment_paths = list()
                for i in range (attachment_file) :
                    attachment_path = input(f"Cho biết đường dẫn file thứ {i + 1}:")
                    attachment_paths.append(attachment_path)
                send_mail(config['MailServer'], int(config['SMTP']), from_mail[1:len(from_mail) - 1], to_email, cc_email, bcc_email
                        , subject, content, attachment_path)
            else :
                send_mail(config['MailServer'], int(config['SMTP']), from_mail[1:len(from_mail) - 1], to_email, cc_email, bcc_email
                        , subject, content, attachments=None)
            print("Đã gửi email thành công")
        elif choice == '2' :
            mailbox()
            # luc nay get mail 
            get_mail(config['MailServer'], int(config['POP3']), username, config['Password'], folder_path, config, filter) 
            while True :
                folder = input("Bạn muốn xem folder nào: ")
                if folder == '0' :
                    break 
                print("Đây là danh sách các email :")
                # duyet qua tat ca cac file trong thu muc 
                count = 1
                file_path = {}
                for file_name in os.listdir(folder_path[temp[folder]]) :
                    file_path[count] = os.path.join(folder_path[temp[folder]], file_name)
                    from_mail2, subject2 = read_msg_file(file_path)
                    print(f"{count}.{from_mail2}, {subject2}")
                    count += 1
                number_file = int(input("Bạn muốn đọc email thứ mấy: ")) 
                print(f"Nội dung email thứ {number_file}: ")
        else :
            break 
    
