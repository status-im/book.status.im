#!/usr/bin/env python3
from __future__ import print_function
import os, re, shutil, stat, io
import git
from datetime import datetime, timezone

# OAuth
import httplib2
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
# Google API
from apiclient import discovery
from apiclient.http import MediaIoBaseDownload
# ODF
import zipfile, xml.dom.minidom
from odf.opendocument import OpenDocumentText, load
from odf import text, teletype
from odf.element import Text
from odf.text import P
from odf.style import Style, TextProperties, ParagraphProperties

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SCOPES = 'https://www.googleapis.com/auth/drive.readonly'
CREDENTIALS='credentials.json'
APPLICATION_NAME = 'Book of Status'
TEXTS_DIR = './texts'
DRIVE_ID = '0AEhafKkWf9UkUk9PVA'


def get_credentials():
    store = Storage(CREDENTIALS)
    credentials = store.get()
    if not credentials or credentials.invalid:
        for filename in os.listdir(os.getcwd()):
            if re.match(r"^client\_secret.+\.json$", filename):
                CLIENT_SECRET_FILE = filename
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
    return credentials


def get_document(service, file_id):
    "Download the Google Doc and return ODF file handle"
    request = service.files().export_media(
        fileId=file_id,
        mimeType='application/vnd.oasis.opendocument.text'
    )
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print(file_id, "downloading... %d%%." % int(status.progress() * 100))
    return fh


def remove_readonly(func, path, _):
    "Clear the readonly bit and reattempt the removal"
    os.chmod(path, stat.S_IWRITE)
    func(path)


def print_e(el, indent=0):
    print(' ' * indent + el.tagName)
    print(' ' * indent + str(el.attributes))
    indent += 1
    for el in el.childNodes[:]:
        print_e(el, indent)


def merge(inputfile, textdoc, document_id):
    style_renaming = {}
    inputtextdoc = load(inputfile)
    # print_e(inputtextdoc.automaticstyles)
    # print('&&&&&&&&&&&')
    # print_e(inputtextdoc.body)
    # return
    # import ipdb; ipdb.set_trace()

    # Need to make a copy of the list because addElement unlinks from the original
    for meta in inputtextdoc.meta.childNodes[:]:
        textdoc.meta.addElement(meta)

    def rename_style(style):
        attr_name = style.getAttribute('name')
        new_attr_name = "%s_doc%s" % (attr_name, document_id)
        print(attr_name, new_attr_name)
        style.setAttribute('name', new_attr_name)
        style_renaming[attr_name] = new_attr_name
        return style
    
    for style in inputtextdoc.styles.childNodes[:]:
        textdoc.styles.addElement(style)

    for autostyle in inputtextdoc.automaticstyles.childNodes[:]:
        textdoc.automaticstyles.addElement(rename_style(autostyle))

    for masterstyles in inputtextdoc.masterstyles.childNodes[:]:
        textdoc.masterstyles.addElement(masterstyles)

    for font in inputtextdoc.fontfacedecls.childNodes[:]:
        textdoc.fontfacedecls.addElement(font)

    for scripts in inputtextdoc.scripts.childNodes[:]:
        textdoc.scripts.addElement(scripts)

    for settings in inputtextdoc.settings.childNodes[:]:
        textdoc.settings.addElement(settings)

    def replace_style(el):

        if not el.attributes:
            return el
        # print(el.attributes)
        stylename = el.attributes.get(('urn:oasis:names:tc:opendocument:xmlns:text:1.0', 'style-name'))
        if stylename and style_renaming.get(stylename):
            print('Replacing style-name: %s with %s' % (stylename, style_renaming.get(stylename)))
            el.attributes[('urn:oasis:names:tc:opendocument:xmlns:text:1.0', 'style-name')] = \
                style_renaming[stylename]

        parent_stylename = el.attributes.get(('urn:oasis:names:tc:opendocument:xmlns:text:1.0', 'parent-style-name'))
        if parent_stylename and style_renaming.get(parent_stylename):
            print('Replacing parent-style-name: %s with: %s' % (parent_stylename, style_renaming.get(parent_stylename)))
            el.attributes[('urn:oasis:names:tc:opendocument:xmlns:text:1.0', 'parent-style-name')] = \
                style_renaming[parent_stylename]

        el.childNodes = [
            replace_style(x) for x in el.childNodes[:]
        ]

        return el

    for body in inputtextdoc.body.childNodes[:]:
        b = replace_style(body)
        textdoc.body.addElement(body)

    textdoc.Pictures = {**textdoc.Pictures, **inputtextdoc.Pictures}
    return textdoc


def replace_tokens(textdoc):
    repo = git.Repo('.')
    assert not repo.bare
    LAST_COMMIT = repo.head.object.hexsha
    GENERATED_TIME = datetime.now(timezone.utc).strftime("%Y.%m.%d %H:%M")

    texts = textdoc.getElementsByType(text.P)
    s = len(texts)
    for i in range(s):
        tmp_text = teletype.extractText(texts[i])
        if '%DATETIME%' in tmp_text or '%LAST_GIT_COMMIT%' in tmp_text:
            tmp_text = tmp_text.replace('%DATETIME%',GENERATED_TIME)
            tmp_text = tmp_text.replace('%LAST_GIT_COMMIT%', LAST_COMMIT)
            new_S = text.P()
            new_S.setAttribute("stylename",texts[i].getAttribute("stylename"))
            new_S.addText(tmp_text)
            texts[i].parentNode.insertBefore(new_S,texts[i])
            texts[i].parentNode.removeChild(texts[i])
    return textdoc


def main():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('drive', 'v3', http=http)
    # TODO Impement pagination
    # TODO Implement recursive directories
    results = service.files().list(
        q="mimeType = 'application/vnd.google-apps.document' and '{0}' in parents".format(DRIVE_ID),  # for subdirectories later
        includeTeamDriveItems=True, corpora='teamDrive', supportsTeamDrives=True, teamDriveId=DRIVE_ID,
        orderBy='createdTime', pageSize=25, fields="nextPageToken, files(id, name)").execute()
    items = results.get('files', [])
    if not items:
        print('No files found.')
    else:
        print('Files:')
        if os.path.exists(TEXTS_DIR):
            # Clear out texts directory so it's easier to see  which documents are changed in git commit
            # (yes I know storing binaries in git is bad)
            shutil.rmtree(TEXTS_DIR, onerror=remove_readonly)
        os.makedirs(TEXTS_DIR)

        # Hacky but to prevent step iterator issue when mergning documents
        doc_stream = get_document(service, items[0]['id'])
        doc_file = os.path.join(TEXTS_DIR, "{0}.odt".format(items[0]['name']))
        with open(doc_file, 'wb') as out:
           out.write(doc_stream.getvalue())
        book_of_status = OpenDocumentText()

        # book_of_status = load(doc_stream)

        # withbreak = Style(name="WithBreak", parentstylename="Standard", family="paragraph")
        # withbreak.addElement(ParagraphProperties(breakbefore="page"))
        # book_of_status.automaticstyles.addElement(withbreak)
        # page_seperator = P(stylename=withbreak,text=u'')

        document_id = 0
        for item in items[:]:
            print('{0} ({1})'.format(item['name'], item['id']))
            doc_stream = get_document(service, item['id'])
            doc_file = os.path.join(TEXTS_DIR, "{0}.odt".format(item['name']))
            with open(doc_file, 'wb') as out:
                out.write(doc_stream.getvalue())

            # TODO unclear if this is the best approach
            # book_of_status.text.addElement(page_seperator)
            book_of_status.text.addElement(text.SoftPageBreak())
            book_of_status = merge(doc_stream, book_of_status, str(document_id))
            document_id += 1

        # Doesn't work well
        book_of_status = replace_tokens(book_of_status)
        book_of_status.save("book-of-status.odt")


if __name__ == '__main__':
    main()
