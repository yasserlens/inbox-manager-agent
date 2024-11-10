import mailbox
import csv
from email import policy
from email.parser import BytesParser
import argparse

def get_body(message):
    if message.is_multipart():
        for part in message.walk():
            if part.is_multipart():
                for subpart in part.walk():
                    if subpart.get_content_type() == 'text/plain':
                        return subpart.get_payload(decode=True)
            elif part.get_content_type() == 'text/plain':
                return part.get_payload(decode=True)
    else:
        return message.get_payload(decode=True)

def mbox_to_csv(mbox_file_path, csv_file_path_template, max_size_mb=1000):
    mbox = mailbox.mbox(mbox_file_path)
    max_size_bytes = max_size_mb * 1024 * 1024

    # First pass: estimate average row size using a sample
    messages = list(mbox)
    sample_size = min(100, len(messages))
    total_sample_size = 0

    for message in messages[:sample_size]:
        body = get_body(message)
        if body:
            body = body.decode('utf-8', errors='replace').replace('\n', ' ').replace('\r', '')
        row = [
            message['subject'],
            message['from'],
            message['date'],
            message['to'],
            message['message-id'],
            body
        ]
        total_sample_size += sum(len(str(field)) for field in row)

    avg_row_size = total_sample_size / sample_size
    rows_per_file = int(max_size_bytes / avg_row_size * 0.9)  # 90% to be safe
    total_parts = (len(messages) + rows_per_file - 1) // rows_per_file

    # Second pass: write files with calculated chunk size
    for part in range(total_parts):
        start_idx = part * rows_per_file
        end_idx = min(start_idx + rows_per_file, len(messages))

        csv_file_path = csv_file_path_template.replace('.csv', f'-part-{part+1}.csv')
        with open(csv_file_path, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(['Subject', 'From', 'Date', 'To', 'Message-ID', 'Body'])

            for message in messages[start_idx:end_idx]:
                body = get_body(message)
                if body:
                    body = body.decode('utf-8', errors='replace').replace('\n', ' ').replace('\r', '')
                else:
                    body = ''
                writer.writerow([
                    message['subject'],
                    message['from'],
                    message['date'],
                    message['to'],
                    message['message-id'],
                    body
                ])

# Usage
if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Convert MBOX to CSV with optional file size limit')
    parser.add_argument('--max-size', type=int, default=1000,
                       help='Maximum size of each CSV file in megabytes (default: 1000)')
    args = parser.parse_args()

    mbox_file_path = 'REPLACE-THIS-WITH-PATH-TO-MBOX-FILE.mbox'
    csv_file_path_template = 'REPLACE-WITH-PATH-TO-CSV-OUTPUT-FILE.csv'
    mbox_to_csv(mbox_file_path, csv_file_path_template, args.max_size)
