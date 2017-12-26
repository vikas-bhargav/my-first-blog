import xlrd
from decimal import Decimal
import email
import imaplib
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from reportlab.pdfgen.canvas import Canvas
from pdfrw import PdfReader
from pdfrw.toreportlab import makerl
from pdfrw.buildxobj import pagexobj


def read_xlsx_file(fileName):
    ExcelFileName= fileName
    workbook = xlrd.open_workbook(ExcelFileName)
    worksheet = workbook.sheet_by_name("Sheet1")
    num_rows = worksheet.nrows  # Number of Rows
    num_cols = worksheet.ncols  # Number of Columns

    result_data_list = list()
    key_list = None

    for curr_row in range(0, num_rows):
        result_data_dict = dict()
        for curr_col in range(0, num_cols, 1):
            if curr_row == 0:
                key = worksheet.cell_value(curr_row, curr_col)
                result_data_dict[key] = ''
            else:
                for i, k in enumerate(key_list):
                    if curr_col == i:
                        result_data_dict[k] = worksheet.cell_value(curr_row, curr_col)

        if key_list is None:
            key_list = result_data_dict.keys()
        if curr_row != 0:
            result_data_list.append(result_data_dict)

    return  result_data_list

# -------- edit pdf -------------- #

def edit_pdf(filePath):
    input_file = filePath
    output_file = filePath
    # Get pages
    reader = PdfReader(input_file)
    pages = [pagexobj(p) for p in reader.pages]
    # Compose new pdf
    canvas = Canvas(output_file)
    for page_num, page in enumerate(pages, start=1):
        # Add page
        canvas.setPageSize((page.BBox[2], page.BBox[3]))
        canvas.doForm(makerl(canvas, page))
        # Draw header
        header_text = "Jhon institute"
        x = 180
        canvas.saveState()
        canvas.setStrokeColorRGB(0, 0, 0)
        canvas.drawImage('input_logo.jpg', height=60, width=110,x=60, y=700)
        canvas.setFont('Times-Roman', 14)
        canvas.drawString(page.BBox[2] -x, 730, header_text)
        # Draw footer
        footer_text = "It’s easy to play any musical instrument: all you have to do is touch the right key at " \
                      "the right time and the instrument will play itself."
        x = 70
        canvas.setStrokeColorRGB(0, 0, 0)
        canvas.setLineWidth(0.5)
        canvas.setFont('Times-Roman', 10)
        canvas.drawString(page.BBox[1] +x, 30, footer_text)
        canvas.restoreState()

        canvas.showPage()

    canvas.save()

# ---------- mail --------------- #

def email_send_get(result_data_list):

    EMAIL_HOST_USER = 'vikas.djangotest@gmail.com'
    EMAIL_HOST_PASSWORD = 'vikas@djangotest'
    EMAIL_HOST_REVEIVER = 'vikasbhargav0@gmail.com'

    # ---- create dir ------- #
    detach_dir = '.'
    if 'attachments' not in os.listdir(detach_dir):
        os.mkdir('attachments')

    userName = EMAIL_HOST_USER
    password = EMAIL_HOST_PASSWORD

    filePath = None
    for data_dict in result_data_list:
        encoded_code = Decimal(data_dict['Encoded Code'])
        try:
            imapSession = imaplib.IMAP4_SSL('imap.gmail.com')
            typ, accountDetails = imapSession.login(userName, password)
            if typ != 'OK':
                print('Not able to sign in!')

            imapSession.select('inbox')
            search_str = '(HEADER Subject ' + str(encoded_code) + ')'

            result, data = imapSession.search(None, search_str)
            if result != 'OK':
                print ('Error searching Inbox.')

            # Iterating over all emails
            for msgId in data[0].split():
                result, messageParts = imapSession.fetch(msgId, '(RFC822)')
                if result != 'OK':
                    print('Error fetching mail.')

                emailBody = messageParts[0][1]
                mail = email.message_from_string((emailBody.decode('utf-8')))
                for part in mail.walk():
                    fileName = part.get_filename()
                    if bool(fileName):
                        filePath = os.path.join(detach_dir, 'attachments', fileName)
                        if not os.path.isfile(filePath) :
                            fp = open(filePath, 'wb')
                            fp.write(part.get_payload(decode=True))
                            fp.close()
                            # edit file
                            edit_pdf(filePath)


            imapSession.close()
            imapSession.logout()
        except :
            print('Not able to download all attachments.')

        # ---- send an email-----#
        msg = MIMEMultipart()
        # storing the senders email address
        msg['From'] = EMAIL_HOST_USER
        # storing the receivers email address
        msg['To'] = EMAIL_HOST_REVEIVER
        # storing the subject
        msg['Subject'] = "Subject of the Mail"
        # string to store the body of the mail
        html = """\
                <html>
                  <head></head>
                  <body>
                    <p>Hey """ + data_dict['User Name'] + """ <br><br><br>

                       Good to see the progress, Please find the attachment of your previous session.
                    </p><br>
                    <p>Thank you, <br>John</p>

                  </body>
                </html>
               """
        body = html
        # attach the body with the msg instance
        msg.attach(MIMEText(body, 'html'))

        # open the file to be sent
        attachment = open(filePath, "rb")
        filename = attachment.name
        # instance of MIMEBase and named as p
        p = MIMEBase('application', 'octet-stream')
        p.set_payload((attachment).read())
        # encode into base64
        encoders.encode_base64(p)
        p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        # attach the instance 'p' to instance 'msg'
        msg.attach(p)

        # creates SMTP session
        s = smtplib.SMTP('smtp.gmail.com', 587)
        s.starttls()
        # Authentication
        s.login(EMAIL_HOST_USER, EMAIL_HOST_PASSWORD)
        text = msg.as_string()
        # sending the mail
        s.sendmail(EMAIL_HOST_USER, EMAIL_HOST_REVEIVER, text)
        # terminating the session
        s.quit()


def main():
    fileName = 'input_Data.xlsx'
    result_data_list = read_xlsx_file(fileName)
    email_send_get(result_data_list)
    print("Task completed successfully!")


if __name__ == '__main__':
    main()