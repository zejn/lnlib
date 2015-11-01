
IBM Notes helper library
========================

This is a Python library for accessing IBM Notes through Windows COM API. The library is tested for reading data from Lotus Notes.

This may serve either as helper or as an example.

Usage
-----

To read data from Notes, you must first establish a NotesSession::

    from lnlib import (get_session, list_views, list_documents,
        describe_document, search, fetch_documents_since)

    session = get_session(ln_key="user's password")

Sometimes session needs to be started in Notes home directory. You can pass ``workdir`` parameter and it will chdir to that directory before initializing session::

    session = get_session(ln_key="user's password", workdir=r'C:\Users\user\AppData\Notes')

With established session, you can now connect to databases::

    server = 'MAIL/ACME'
    nsfpath = r'mail\nikolatesla.nsf'

    db = session.GetDatabase(server, nsfpath)

If database is local nsf file, server is empty and nsfpath is path to local file::

    server = ''
    nsfpath = r'Users\user\AppData\Notes\mail.nsf'

    db = session.GetDatabase(server, nsfpath)

Having a database handle, you can now explore available views and get specific one::

    print(db.Views)
    inbox = db.GetView('$Inbox')

Get first document from view::

    doc = inbox.GetFirstDocument()

Print its UNID and NotesURL::

    print(doc.UniversalID)
    print(doc.NotesURL)

..

 * You can use NotesURL in web page and browser will call Notes to open document.

Describe document and convert it to pure python object::

    describe_document(doc)

    json_doc = lndoc2obj(doc)
    print(json_doc)

Iterate over documents in view::

    for doc in list_documents(inbox):
        print(doc.UniversalID)

Search for text in a view::

    for doc in search(inbox, "Edison"):
        print(doc.UniversalID)

Fetch only documents, changed in last 24 hours::

    since = datetime.datetime.now() - datetime.timedelta(1)
    for doc in fetch_documents_since(session, db, since):
        print(doc.UniversalID)

..

 * Using ``fetch_documents_since``, one can keep up with document changes without scanning full database each turn.
