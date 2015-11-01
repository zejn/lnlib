import datetime
import json
import os
import re

import pywintypes
from win32com.client import Dispatch


class ItemType:
    ACTIONCD = 16
    ASSISTANTINFO = 17
    ATTACHMENT = 1084
    AUTHORS = 1076
    COLLATION = 2
    DATETIMES = 1024
    EMBEDDEDOBJECT = 1090
    ERRORITEM = 256
    FORMULA = 1536
    HTML = 21
    ICON = 6
    LSOBJECT = 20
    MIME_PART = 25
    NAMES = 1074
    NOTELINKS = 7
    NOTEREFS = 4
    NUMBERS = 768
    OTHEROBJECT = 1085
    QUERYCD = 15
    READERS = 1075
    RFC822Text = 1282
    RICHTEXT = 1
    SIGNATURE = 8
    TEXT = 1280
    UNAVAILABLE = 512
    UNKNOWN = 0
    USERDATA = 14
    USERID = 1792
    VIEWMAPDATA = 18
    VIEWMAPLAYOUT = 19

ITEM_TYPES = dict([(v, k) for k, v in ItemType.__dict__.items() if not k.startswith('_')])


def get_session(ln_key, workdir=None):
    "helper function for creating a new NotesSession"
    if workdir:
        os.chdir(workdir)
    session = Dispatch('Lotus.NotesSession')
    session.Initialize(ln_key)
    return session


def list_views(db):
    "list views in database"
    for v in db.Views:
        yield v.Name, v


def list_documents(view):
    "a helper function for listing views easily in for loops"
    doc = view.GetFirstDocument()

    while doc:
        yield doc
        doc = view.GetNextDocument(doc)


def search(view, text):
    "full text search on a given view"
    num = view.FTSearch(text)
    print 'Najdenih zadev:', num

    doc = view.GetFirstDocument()
    while doc:
        yield doc
        doc = view.GetNextDocument()


def windt2datetime(pytim):
    "convert windows time object (PyTime) to datetime"
    now_datetime = datetime.datetime(
        year=pytim.year,
        month=pytim.month,
        day=pytim.day,
        hour=pytim.hour,
        minute=pytim.minute,
        second=pytim.second
    )
    return now_datetime


def _dt(obj):
    "convert pywintypes.Time to datetime"
    if not type(obj) == type(pywintypes.Time(1)):
        return obj
    else:
        return windt2datetime(obj)


def lndoc2obj(doc):
    """
    Convert NotesDocument to a Python dictionary-like structure.

    Structure:

        {
            '<first_item_name>': {
                'values': ('<first value>', '<second value>),
                'type': 'TEXT',
                'last_modified': datetime.datetime(2015, 2, 10, 12, 45, 34)
            },
            '<second_item_name>': {
                'values': (datetime.datetime(2015, 2, 1, 0, 0, 0),),
                'type': 'DATETIMES',
                'last_modified': datetime.datetime(2015, 2, 2, 10, 13, 43)
            }
        }
    """
    vals = []
    if doc.Items:
        for k in doc.Items:
            itm = doc.GetFirstItem(k.Name)
            if itm.Type in (ItemType.DATETIMES, ItemType.RFC822Text):
                val = tuple([_dt(i) for i in itm.Values])
            else:
                val = doc.GetItemValue(k.Name)

            r = (k.Name, {
                'values': val,
                'type': ITEM_TYPES[itm.Type],
                'last_modified': _dt(itm.LastModified),
            })
            vals.append(r)
    return dict(vals)


def fetch_documents_since(session, db, since):
    """
    Query NotesDatabase by document modification time, and fetch documents.

    `session` is a NotesSession instance,

    `db` is a NotesDatabase instance to fetch documents from,

    `since` is a datetime.datetime, to only fetch documents updated after this time.
    If None, all documents are returned.
    """
    if since is None:
        docs = db.AllDocuments
    elif isinstance(since, datetime.datetime):
        from_lntime = session.CreateDateTime(since.strftime('%d/%m/%Y %H:%M:%S'))
        docs = db.GetModifiedDocuments(from_lntime)
    return docs


def describe_view(view):
    "describe view and first document"
    doc = view.GetFirstDocument()
    print 'View:', view.Name
    describe_document(doc)


def describe_document(doc):
    "describe document"
    items = []
    for k in doc.Items:
        itm = doc.GetFirstItem(k.Name)
        if itm.Type == ItemType.DATETIMES:
            val = tuple([_dt(i) for i in doc.GetItemValue(itm.Name)])
        else:
            val = doc.GetItemValue(itm.Name)
        r = (itm.Name, ITEM_TYPES[itm.Type], val)
        # r = (k.Name, [_dt(i) for i in doc.GetItemValue(k.Name)])
        items.append(r)
    items.sort()
    for i in items:
        print i

def note_created(unid):
    "get creation time from a Note's UniversalID"
    assert re.match('^[0-9A-F]{32}$', unid.strip().upper()), 'not a Note ID'
    return hextimedate2datetime(unid[16:])

def datetime2jdn(dt):
    "converts python datetime to julian date number"
    "calculation follows https://en.wikipedia.org/wiki/Julian_day"

    a = (14 - dt.month) // 12
    y = dt.year + 4800 - a
    m = dt.month + 12 * a - 3

    return dt.day + (153 * m + 2) // 5 + 365 * y + y // 4 - y // 100 + y // 400 - 32045


def jdn2datetime(jdn):
    "converts Julian date number to python datetime"

    f = jdn + 1401 + (((4 * jdn + 274277) / 146097) * 3) / 4 - 38
    e = 4 * f + 3
    g = (e % 1461) / 4
    h = 5 * g + 2
    D = (h % 153) / 5 + 1
    M = (h / 153 + 2) % 12 + 1
    Y = e / 1461 - 4716 + (12 + 2 - M) / 12

    return datetime.datetime(Y, M, D)


def hextimedate2datetime(td):
    """
    extract date from Lotus Notes ID

    Note's UniversalID contains a creation time encoded in hex as a
    Julian date number, hundredths of second and time zone information
    
    More information on what a Notes ID is composed of is available at
    http://www-12.lotus.com/ldd/doc/domino_notes/9.0/api90ug.nsf/85255d56004d2bfd85255b1800631684/00d000c1005800c985255e0e00726863?OpenDocument
    or if the url is broken, see IBM Notes C API User Guide, Appendix 1, "Anatomy of a Note ID".
    """
    assert len(td) == 16

    dt = jdn2datetime(int(td[2:8], 16))
    tim = int(td[8:], 16)  # in 1/100th of a second
    dt2 = dt.replace(hour=tim // 360000, minute=tim // 6000 % 60, second=tim // 100 % 60, microsecond=(tim % 100) * 10000)

    flags = int(td[:2], 16)
    if flags & 1 << 7:
        if flags & 1 << 6:
            prefix = +1
        else:
            prefix = -1

        hours = flags & 15
        minutes = (flags >> 4 & 3) * 15
        dt2 = dt2 + datetime.timedelta(seconds=prefix * (hours * 3600 + minutes * 50))

    return dt2


class DateTimeEncoder(json.JSONEncoder):
    """
    JSON encoder class with support for encoding datetime
    """
    def default(self, obj):
        if isinstance(obj, datetime.datetime):
            return obj.isoformat()
        elif isinstance(obj, datetime.date):
            return obj.isoformat()
        return json.JSONEncoder.default(obj)


def as_json(obj):
    "helper json dumps function with support for datetime"
    return json.dumps(obj, cls=DateTimeEncoder, ensure_ascii=True, sort_keys=True)

if __name__ == "__main__":
    import unittest

    class JDNTest(unittest.TestCase):
        def runTest(self):
            d = datetime.datetime(2015, 1, 29)

            jdn = datetime2jdn(d)
            self.assertEqual(2457052, jdn)

            d2 = jdn2datetime(jdn)
            self.assertEqual(d, d2)

    class HexTimeDateTest(unittest.TestCase):
        def runTest(self):
            unid = 'C1257DDC0028EDC8'
            unid = 'C1257DDC002B4524'
            unid = 'C1257DDC00481076'
            created = hextimedate2datetime(unid)
            self.assertEqual(created, datetime.datetime(2015, 1, 29, 14, 7, 8, 60000))
            unid = 'C1257DDC00486761'
            created = hextimedate2datetime(unid)
            self.assertEqual(created, datetime.datetime(2015, 1, 29, 14, 10, 50, 570000))
                

            unid = '76089E66C2532CE0C125647A0030ED69'
            created = note_created(unid)
            self.assertEqual(created, datetime.datetime(1997, 4, 15, 9, 54, 25, 50000))


    unittest.main()
