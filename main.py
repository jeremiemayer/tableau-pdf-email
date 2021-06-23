import smtplib
import tableauserverclient as TSC
import logging
import functools
import tempfile
import shutil
import PyPDF2
import os.path
import calendar
from datetime import datetime, timedelta, date
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate

def get_recipients(file):
    recipients = open(file,'r').readlines()
    recipients = [r.strip() for r in recipients] #remove new lines and whitespace
    return recipients

def get_filter(type):
    now = datetime.now() - timedelta(days=1)
    month = now.month

    if month >= 6:
        if(type=='tableau'):
            return 'Filter1:Value'
        elif(type=='email'):
            return 'Filter: Value'
    else:
        if(type=='tableau'):
            return 'Filter1:Value2'
        elif(type=='email'):
            return 'Filter: Value2'

def get_views_for_workbook(server, workbook_id):
    workbook = server.workbooks.get_by_id(workbook_id)
    server.workbooks.populate_views(workbook)

    pdf_views = []
    for view in workbook.views:
        if view.name in ('Tab 1','Tab 2'):
            pdf_views.append(view)
    return pdf_views

def download_pdf(server, tempdir, view):
    logging.info("Exporting {}".format(view.id))
    destination_filename = os.path.join(tempdir, view.id)

    # setup filters
    option_factory = getattr(TSC, 'PDFRequestOptions')
    options = option_factory().vf(*get_filter('tableau').split(':'))

    server.views.populate_pdf(view,options)
    with open(destination_filename, 'wb') as f:
        f.write(view.pdf)

    return destination_filename

def combine_into(dest_pdf, filename):
    dest_pdf.append(filename)
    return dest_pdf

def cleanup(tempdir):
    shutil.rmtree(tempdir)

def send_mail(send_from, send_to, subject, text, files=None, server="local.smtp.com:2525"):
    assert isinstance(send_to, list)

    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['Bcc'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(text))

    for f in files or []:
        with open(f, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(f)
            )
        # After the file is closed
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
        msg.attach(part)


    smtp = smtplib.SMTP(server)
    smtp.send_message(msg)
    #smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.close()

def main():

    logging.basicConfig(level='ERROR')
    tempdir = tempfile.mkdtemp('temp')
    logging.debug("Saving to tempdir: %s", tempdir)

    tableau_auth = TSC.TableauAuth('user','pwd')
    server = TSC.Server('https://reports2.agconnect.org', use_server_version=True)

    project_name = 'project'
    workbook_name = 'workbook'
    recipients_file = 'recipients.txt'
    file_name = 'output.pdf'
    subject = 'Tableau PDF'
    body = '''Attached is a summary of weekly progress.
    
Notes:
* This report includes only Canadian data
* {filters}
'''.format(filters=get_filter('email'))

    with server.auth.sign_in(tableau_auth):

        options = TSC.RequestOptions()
        options.filter.add(TSC.Filter(TSC.RequestOptions.Field.ProjectName,TSC.RequestOptions.Operator.Equals, project_name))
        options.filter.add(TSC.Filter(TSC.RequestOptions.Field.Name,TSC.RequestOptions.Operator.Equals, workbook_name))

        filtered_workbooks, pagination_item = server.workbooks.get(req_options=options)

        # Download workbook to a temp directory
        if len(filtered_workbooks) == 0:
            logging.debug('No workbook named {} found.'.format('workbook'))
        else:
            workbook_id = filtered_workbooks[0].id

            get_list = functools.partial(get_views_for_workbook, server)
            download = functools.partial(download_pdf, server, tempdir)

            downloaded = (download(x) for x in get_list(workbook_id))
            output = functools.reduce(combine_into, downloaded, PyPDF2.PdfFileMerger())
                
            # save file
            with open(file_name, 'wb') as f:
                output.write(f);            

            # scale pdf
            pdf = PyPDF2.PdfFileReader(file_name)
            page1 = pdf.getPage(0)
            page1.mediaBox.lowerLeft = (20,220)
            page1.mediaBox.lowerRight = (320,220)
            page1.mediaBox.upperLeft = (20,745)
            page1.mediaBox.upperRight = (320,745)
            page1.scaleBy(3.5)

            page2 = pdf.getPage(1)
            page2.mediaBox.lowerLeft = (25,0)
            page2.mediaBox.lowerRight = (235,0)
            page2.mediaBox.upperLeft = (25,745)
            page2.mediaBox.upperRight = (235,745)
            page2.scaleBy(5)
            
            writer = PyPDF2.PdfFileWriter()
            writer.addPage(page1)
            writer.addPage(page2)
            
            with open(file_name,'wb+') as f:
                writer.write(f)

    
    send_mail("test@test.ca",get_recipients(recipients_file),subject,body,[file_name])

if __name__ == '__main__':
    main()