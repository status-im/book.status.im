#!/usr/bin/env python3
from __future__ import print_function
import os, re, shutil, stat, io
import git

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
from odf.element import Text
from odf.text import P

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
    request = service.files().export_media(fileId=file_id,
                                                 mimeType='application/vnd.oasis.opendocument.text')
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

def merge(inputfile, textdoc):
    # TODO replace 
    inputtextdoc = load(inputfile)
    # Need to make a copy of the list because addElement unlinks from the original
    for meta in inputtextdoc.meta.childNodes[:]:
        textdoc.meta.addElement(meta)

    for font in inputtextdoc.fontfacedecls.childNodes[:]:
        textdoc.fontfacedecls.addElement(font)

    for style in inputtextdoc.styles.childNodes[:]:
        textdoc.styles.addElement(style)

    for autostyle in inputtextdoc.automaticstyles.childNodes[:]:
        textdoc.automaticstyles.addElement(autostyle)

    for scripts in inputtextdoc.scripts.childNodes[:]:
        textdoc.scripts.addElement(scripts)

    for settings in inputtextdoc.settings.childNodes[:]:
        textdoc.settings.addElement(settings)

    for masterstyles in inputtextdoc.masterstyles.childNodes[:]:
        textdoc.masterstyles.addElement(masterstyles)

    for body in inputtextdoc.body.childNodes[:]:
        textdoc.body.addElement(body)

    textdoc.Pictures = inputtextdoc.Pictures
    return textdoc

def main():
    repo = git.Repo('.')
    assert not repo.bare
    sha = repo.head.object.hexsha
    exit(sha)

    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('drive', 'v3', http=http)
    # TODO Impement pagination
    # TODO Implement recursive directories
    results = service.files().list(
        q= "mimeType = 'application/vnd.google-apps.document' and '{0}' in parents".format(DRIVE_ID),  # for subdirectories later
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

        book_of_status = OpenDocumentText()
        for item in items:
            print('{0} ({1})'.format(item['name'], item['id']))
            doc_stream = get_document(service, item['id'])
            doc_file = os.path.join(TEXTS_DIR, "{0}.odt".format(item['name']))

            with open(doc_file, 'wb') as out:
                out.write(doc_stream.getvalue())

            book_of_status = merge(doc_stream, book_of_status)
        book_of_status.save("book-of-status.odt")


if __name__ == '__main__':
    main()
