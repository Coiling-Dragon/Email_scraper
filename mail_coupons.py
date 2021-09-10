import imaplib
import email
from email.header import decode_header
import pandas as pd


mails_df = pd.read_csv('mails.csv')
csv_values = mails_df.values

c = 1

with open('mails_with_coupons.csv', 'w', encoding='utf-8') as f:
    out_row = 'EMAIL,PASS,COUPONS\n'
    f.write(out_row)

for each in csv_values:
    user = each[0]
    password = each[1]

# Mailbox interaction

    M = imaplib.IMAP4_SSL('imap.mail.com')
    M.login(user, password)
    M.select('Inbox')
    typ, data = M.search(None, 'ALL')
    ids = data[0]
    id_list = ids.split()
    # get the most recent email id
    latest_email_id = int(id_list[-1])

    COUPON_AMOUNT = '15'

    # iterate through 15 messages in descending order starting with latest_email_id
    # the '-1' dictates reverse looping order
    for i in range(latest_email_id, latest_email_id - 15, -1):
        typ, data = M.fetch(str(i), '(RFC822)')

        for response_part in data:
            if isinstance(response_part, tuple):
                mail_bytes = response_part[1].decode('UTF-8')
                msg = email.message_from_string(mail_bytes)

                varSubject = msg['subject']
                varFrom = msg['from']
                varSubject = decode_header(varSubject)[0][0]

                if f'$coupon' in str(varSubject):
                    print(f'{c} Mail: {user}\n  Subject: {varSubject}\n')
                    with open('mails_with_coupons.csv', 'a') as f:
                        row = f'{user},{password},"${COUPON_AMOUNT}"\n'
                        f.write(row)
                    c += 1

data_frame = pd.read_csv('mails_with_coupons.csv', encoding="utf-8").drop_duplicates(
    subset='EMAIL', keep='first', inplace=False)
data_frame.to_csv('mails_with_coupons.csv', index=False, encoding="utf-8")
