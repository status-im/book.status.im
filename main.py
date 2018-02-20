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
from odf import office, text, teletype
from odf.element import Text
from odf.text import P
from odf.style import Style, TextProperties, ParagraphProperties

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SCOPES = 'https://www.googleapis.com/auth/drive.readonly'
CREDENTIALS = 'credentials.json'
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


def rename_style(style, style_renaming, document_id):
    if style.tagName == 'style:style':
        # Rename style, and parent style.
        for style_type in ['name', 'parentstylename']:
            attr_name = style.getAttribute(style_type)
            if attr_name:
                new_attr_name = "%s_doc%s" % (attr_name, document_id)
                style.setAttribute(style_type, new_attr_name)
                style_renaming[attr_name] = new_attr_name
    return style


def replace_style(el, style_renaming, document_id):

    if el.attributes:

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
        replace_style(x, style_renaming, document_id) for x in el.childNodes[:]
    ]

    return el


def merge(inputfile, textdoc, document_id):
    style_renaming = {}
    inputtextdoc = load(inputfile)

    # Need to make a copy of the list because addElement unlinks from the original
    for meta in inputtextdoc.meta.childNodes[:]:
        textdoc.meta.addElement(meta)

    for autostyle in inputtextdoc.automaticstyles.childNodes[:]:
        s = rename_style(autostyle, style_renaming, document_id)
        textdoc.automaticstyles.addElement(s)

    for style in inputtextdoc.styles.childNodes[:]:
        s = rename_style(style, style_renaming, document_id)
        textdoc.styles.addElement(style)

    for masterstyles in inputtextdoc.masterstyles.childNodes[:]:
        textdoc.masterstyles.addElement(masterstyles)

    for font in inputtextdoc.fontfacedecls.childNodes[:]:
        textdoc.fontfacedecls.addElement(font)

    for scripts in inputtextdoc.scripts.childNodes[:]:
        textdoc.scripts.addElement(scripts)

    for settings in inputtextdoc.settings.childNodes[:]:
        textdoc.settings.addElement(settings)

    for body in inputtextdoc.body.childNodes[:]:
        b = replace_style(body, style_renaming, document_id)
        textdoc.body.addElement(b)

    textdoc.Pictures.update(inputtextdoc.Pictures)

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
            tmp_text = tmp_text.replace('%DATETIME%', GENERATED_TIME)
            tmp_text = tmp_text.replace('%LAST_GIT_COMMIT%', LAST_COMMIT)
            new_S = text.P()
            new_S.setAttribute("stylename", texts[i].getAttribute("stylename"))
            new_S.addText(tmp_text)
            texts[i].parentNode.insertBefore(new_S, texts[i])
            texts[i].parentNode.removeChild(texts[i])
    return textdoc


def main():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('drive', 'v3', http=http)
    # TODO Impement pagination
    folder_results = service.files().list(
        q="mimeType = 'application/vnd.google-apps.folder' and '{0}' in parents".format(DRIVE_ID),
        includeTeamDriveItems=True, corpora='teamDrive', supportsTeamDrives=True, teamDriveId=DRIVE_ID,
    ).execute()
    folders = folder_results.get('files', [])

    for folder in folders:
        results = service.files().list(
            q="mimeType = 'application/vnd.google-apps.document' and '{0}' in parents".format(folder['id']),
            includeTeamDriveItems=True, corpora='teamDrive', supportsTeamDrives=True, teamDriveId=DRIVE_ID,
            orderBy='createdTime', pageSize=25, fields="nextPageToken, files(id, name)"
        ).execute()
        files = results.get('files', [])

        if not files:
            print('No files found.')
        else:
            print('Files:')
            if os.path.exists(TEXTS_DIR):
                # Clear out texts directory so it's easier to see  which documents are changed in git commit
                # (yes I know storing binaries in git is bad)
                shutil.rmtree(TEXTS_DIR, onerror=remove_readonly)
            os.makedirs(TEXTS_DIR)

            document_id = 0
            for doc_file in files[:]:

                print('{0} ({1})'.format(doc_file['name'], doc_file['id']))
                doc_stream = get_document(service, doc_file['id'])
                doc_file = os.path.join(TEXTS_DIR, "{0}.odt".format(doc_file['name']))
                with open(doc_file, 'wb') as out:
                    out.write(doc_stream.getvalue())

                # OpenDocumentText() can not be trusted to be compatible with
                #the WebODF viewer rendering, causing 'Step iterator must be on a step' error.
                # Instead use the first document as our source.
                if document_id == 0:
                    guide = load(doc_stream)
                else:
                    guide = merge(doc_stream, guide, str(document_id))
                document_id += 1

            guide = replace_tokens(guide)
            guide.save("{}.odt".format(folder["name"]))


if __name__ == '__main__':
    main()
