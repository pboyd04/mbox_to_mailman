import argparse
import mailbox
import datetime
import dateutil.parser
import zipfile
from sortedcontainers import SortedDict

headers = ['From', 'Date', 'Subject', 'In-Reply-To', 'References', 'Message-Id', 'Message-ID']

def part_to_text(part, dateStr):
    contentType = part.get_content_type()
    if contentType == 'multipart/alternative' or contentType == 'multipart/related':
        return None
    if contentType == 'text/html':
        return None 
    if contentType != 'text/plain':
        attachmentName = part.get_filename()
        if attachmentName == None:
            return None
        with zipfile.ZipFile(dateStr+'-attachments.zip', 'a') as myzip:
            myzip.writestr(attachmentName, part.get_payload(decode=True))
        return None
    charset = part.get_content_charset()
    if not charset:
        return None
    if charset == 'x-unknown' or charset == 'cp-850':
        return None
    text = str(part.get_payload(decode=True), encoding=charset, errors='ignore')
    try:
        text = str(text.encode('ascii'), 'ascii')
    except UnicodeEncodeError:
        return None
    except UnicodeDecodeError:
        return None
    return text

def message_to_text(msg, dateStr):
    payload = msg.get_payload()
    if isinstance(payload, str):
        return payload
    text = ''
    for h in msg.items():
        text += h[0]+': '+str(h[1])+'\n'
    for part in msg.walk():
        part = part_to_text(part, dateStr)
        if part:
            text += part
    return text

def mailbox_parse(mb):
    messageDB = {}
    for message in mb:
        for header in message.keys():
            if header not in headers:
                del message[header]
        try:
            date = datetime.datetime.strptime(message['Date'], "%a, %d %b %Y %X %z")
        except ValueError as e:
            date = dateutil.parser.parse(message['Date'], fuzzy=True)
        dateStr = date.strftime('%Y-%m')
        if date.year not in messageDB:
            messageDB[date.year] = {}
        if date.month not in messageDB[date.year]:
            messageDB[date.year][date.month] = SortedDict({})
        text = message_to_text(message, dateStr)
        messageDB[date.year][date.month][date.timestamp()] = text;
    return messageDB


def main():
    parser = argparse.ArgumentParser(description='Convert mbox to mailbox style archive.')
    parser.add_argument('mbox_file', help='.mbox file to parse')
    args = parser.parse_args()

    mb = mailbox.mbox(args.mbox_file, create=False)
    messageDB = mailbox_parse(mb)
    for year in messageDB:
        for month in messageDB[year]:
            messages = messageDB[year][month]
            fileObj = open(str(year)+'-'+str(month).zfill(2)+'.txt', 'w')
            for text in messages.values():
                fileObj.write(text)
                fileObj.write('\n')
            fileObj.close()


if __name__ == '__main__':
    main()
