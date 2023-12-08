import socket
import base64
import os
import random
import string
from email import policy
from email.parser import BytesParser
import re
import threading
import time
import json



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
        if msg['Cc'] :
            print(f"Cc: {msg['Cc']}")
        # In nội dung email
        print("\nBody:")
        body2 = str(msg.get_body().get_content())
        print(body2)

        # # Tạo thư mục để lưu đính kèm (nếu chưa tồn tại)
        # os.makedirs(output_folder, exist_ok=True)

        # Xử lý đính kèm nếu có
        for idx, part in enumerate(msg.iter_parts()):
            if part.get_content_disposition() == 'attachment':
                number = int(input("Trong email này có attached file, bạn có muốn save không(1: có , other: không): "))
                if number == 1 :
                    output_folder = input("Cho biết đường dẫn bạn muốn lưu: ")
                    output_folder = output_folder[1 : len(output_folder) - 1]
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
        file_name = os.path.basename(msg_file_path)
        file_path = os.path.join(".\\Seen", file_name)
        with open(file_path, 'wb') as file :
            file.write("da xem".encode())
        return from_mail2, subject2, body2
    except Exception as e:
        print(f"An error occurred: {e}")



read_msg_file("C:\\Users\\ASUS ZenBook\\OneDrive\\Máy tính\\python\\Project-Socket\\Inbox\\20231208110055715.msg")