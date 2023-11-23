
import extract_msg

msg_file = "C:\\Users\\ASUS ZenBook\\OneDrive\\Máy tính\\python\\Project-Socket\\20231123150045902.msg"
msg = extract_msg.Message(msg_file)

msg_sender = msg.sender
msg_date = msg.date
msg_subject = msg.subject
msg_body = msg.body

print(f'Sender: {msg_sender}')
print(f'Sent On: {msg_date}')
print(f'Subject: {msg_subject}')
print(f'Body: {msg_body}')

msg.close()

# Example usage
