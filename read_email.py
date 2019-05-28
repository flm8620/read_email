from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr, getaddresses, parsedate
from email import message_from_bytes, message_from_string
import os
import time
import xlsxwriter

# indent用于缩进显示:
def extract_info(msg, result, f=None, indent=0):
    if indent == 0:
        for header in msg.keys():#['From', 'To', 'Subject']:
            value = msg.get(header, '')
            contents = []
            if value:
                if header=='Date':
                    contents = decode_str(value)
                    t = time.localtime(time.mktime(parsedate(contents[0])))
                    result['Date'] = '%d/%d/%d' % (t.tm_year, t.tm_mon, t.tm_mday)
                elif header=='From' or header=='To' or header=='Cc':
                    for (hdr, addr) in getaddresses([value]):
                        name = decode_str(hdr)[0]
                        contents.append(u'%s <%s>' % (name, addr))
                    result[header] = contents
                elif header=='Subject':
                    contents = decode_str(value)
                    result['Subject'] = contents[0]
                else:
                    contents = decode_str(value)
                    
                f.write('%s%s:\n' % ('  ' * indent, header))
                for v in contents:        
                    f.write('%s%s\n' % ('  '*(indent+1), v))
                f.write('')
    if (msg.is_multipart()):
        parts = msg.get_payload()
        for n, part in enumerate(parts):
            f.write('%sPart %s/%s\n'              % ('  ' * indent, n+1, len(parts)))
            f.write('%s{\n' % ('  ' * indent))
            extract_info(part, result, f, indent + 1)
            f.write('%s}\n' % ('  ' * indent))
    else:
        content_type = msg.get_content_type()
        f.write('%sContentDisposition: %s'%('  '*indent, msg.get_content_disposition()))
        if content_type=='text/plain' or content_type=='text/html':
            content = msg.get_payload(decode=True)
            charset = guess_charset(msg)
            if charset:
                content = content.decode(charset)
            f.write('%sText:\n'       % ('  ' * indent))
            f.write(content)
            if content_type=='text/plain':
                result['Text'] = content
        else:
            filename = decode_str(msg.get_filename('unknown_file_name'))[0]
            f.write('%sAttachment: %s, %s\n' % ('  ' * indent, content_type, filename))
            if msg.get_content_disposition() == 'attachment':
                result['Attachments'].append(filename)        

def decode_str(s):
    r = []
    for value, charset in decode_header(s):
        r.append(value.decode(charset) if charset else value)
    return r


def guess_charset(msg):
    charset = msg.get_charset()
    if charset is None:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset

workbook = xlsxwriter.Workbook('output.xlsx')
ws = workbook.add_worksheet()
fmt = workbook.add_format({'text_wrap': True, 'valign': 'top'})
row = 0

ws.write(row, 0, '时间')
ws.set_column(0, 0, 10)
ws.write(row, 1, '发件人')
ws.set_column(1, 1, 40)
ws.write(row, 2, '收件人')
ws.set_column(2, 2, 40)
ws.write(row, 3, '抄送')
ws.set_column(3, 3, 40)
ws.write(row, 4, '主题')
ws.set_column(4, 4, 50)
ws.write(row, 5, '正文')
ws.set_column(5, 5, 50)
ws.write(row, 6, '附件')
ws.set_column(6, 6, 70)
row+=1
folder = 'emls'
for path in os.listdir(folder): 
    file = folder + '/' + path
    if os.path.isfile(file) and file.endswith('.eml'):
        print('Reading %s' % file)
        f = open(file, "r", encoding='utf-8')
        email_string = f.read()
        f.close()
        out = open(os.path.basename(path)+".txt", "w", encoding='utf-8')
        #out = open(os.devnull, "w", encoding='utf-8')
        msg = message_from_string(email_string)
        result = {}
        result['Attachments'] = []
        result['Text'] = ''
        result['Cc'] = []
        extract_info(msg, result, out)
        out.close()
        
        ws.write(row, 0, result['Date'])
        ws.write(row, 1, '\r\n'.join(result['From']), fmt)
        ws.write(row, 2, '\r\n'.join(result['To']), fmt)
        ws.write(row, 3, '\r\n'.join(result['Cc']), fmt)
        ws.write(row, 4, result['Subject'], fmt)
        
        text = result['Text']
        text = text.replace("\r\n", "\n")
        text = text.split('\n')
        text_only_this_email = []
        for line in text:
            if line.startswith('发件人：') or line.startswith('From:'):
                break
            text_only_this_email.append(line)
            
        ws.write(row, 5, '\r\n'.join(text_only_this_email), fmt)
        ws.write(row, 6, '\r\n'.join(result['Attachments']), fmt)
        row+=1


workbook.close()
