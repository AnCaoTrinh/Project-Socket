import os
import uuid
from email import policy
from email.parser import BytesParser
import re
def read_msg_file(msg_file_path, output_folder='attachments'):
    try:
        # Đọc nội dung của file .msg
        with open(msg_file_path, 'rb') as file:
            msg_content = file.read()

        # Giải mã base64 và phân tích cú pháp
        msg = BytesParser(policy=policy.default).parsebytes(msg_content)

        # In thông tin email
        print(f"Subject: {msg['Subject']}")
        print(f"From: {msg['From']}")
        print(f"To: {msg['To']}")
        print(f"Cc: {msg['Cc']}")
        print(f"Bcc: {msg['Bcc']}")

        # In nội dung email
        print("\nBody:")
        print(msg.get_body().get_content())

        # Tạo thư mục để lưu đính kèm (nếu chưa tồn tại)
        os.makedirs(output_folder, exist_ok=True)

        # Xử lý đính kèm nếu có
        for idx, part in enumerate(msg.iter_parts()):
            if part.get_content_disposition() == 'attachment':
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

    except Exception as e:
        print(f"An error occurred: {e}")

# Đường dẫn đến file .msg cụ thể
msg_file_path = 'C:\\Users\\ASUS ZenBook\\OneDrive\\Máy tính\\python\\Project-Socket\\20231122132827696.msg'

# Thư mục để lưu đính kèm
output_folder = 'C:\\Users\\ASUS ZenBook\\OneDrive\\Máy tính\\python\\Project-Socket'

# Gọi hàm để đọc và xử lý nội dung
read_msg_file(msg_file_path, output_folder)