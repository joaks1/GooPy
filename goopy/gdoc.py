#! /usr/bin/env python

import os
import sys
import re
import gdata.docs
import gdata.docs.service
import gdata.spreadsheet
import gdata.spreadsheet.service
from gdata.service import RequestError

from goopy.utils.errors import (SpreadsheetNameError, WorksheetNameError,
        GoogleSpreadsheetClientKeyError, GoogleSpreadsheetClientIdError)
from goopy.utils.messaging import get_logger

_LOG = get_logger(__name__, 'INFO')

class GoogleSpreadsheetClient(object):
    
    def __init__(self, username, password,
            spreadsheet_name_or_key = None,
            worksheet_name_or_id = 'od6'):
        self.doc_client = gdata.docs.service.DocsService()
        self.spreadsheet_client = gdata.spreadsheet.service.SpreadsheetsService()
        self.doc_client.email = username
        self.doc_client.password = password
        self.spreadsheet_client.email = username
        self.spreadsheet_client.password = password
        self.doc_client.ProgrammaticLogin()
        self.spreadsheet_client.ProgrammaticLogin()
        self._spreadsheet_key = ''
        self._worksheet_id = worksheet_name_or_id
        if spreadsheet_name_or_key:
            self.spreadsheet_key = spreadsheet_name_or_key
            self.worksheet_id = worksheet_name_or_id
        
    def _get_spreadsheet_key(self):
        return self._spreadsheet_key

    def _set_spreadsheet_key(self, spreadsheet_name_or_key):
        try:
            self.spreadsheet_client.GetSpreadsheetsFeed(
                    key=spreadsheet_name_or_key)
            key = spreadsheet_name_or_key
        except RequestError, key_error:
            try:
                key = self._get_spreadsheet_key_by_name(spreadsheet_name_or_key)
            except SpreadsheetNameError, name_error:
                raise GoogleSpreadsheetClientKeyError(
                    '\n  Problem finding spreadsheet by the name or key {0!r}\n'
                    'Trying {0!r} as the key resulted in the following '
                    'RequestError:\n  {1}\n'
                    'Trying {0!r} as the name resulted in the following '
                    'SpreadSheetNameError:\n  {2}'.format(
                            spreadsheet_name_or_key,
                            key_error,
                            name_error))
            pass
        self._spreadsheet_key = key
                    
    spreadsheet_key = property(_get_spreadsheet_key, _set_spreadsheet_key)

    def _get_worksheet_id(self):
        return self._worksheet_id

    def _set_worksheet_id(self, worksheet_name_or_id):
        try:
            self.spreadsheet_client.GetWorksheetsFeed(
                    key = self.spreadsheet_key,
                    wksht_id = worksheet_name_or_id)
            wid = worksheet_name_or_id
        except RequestError, id_error:
            try:
                wid = self._get_worksheet_id_by_name(worksheet_name_or_id)
            except WorksheetNameError, name_error:
                raise GoogleSpreadsheetClientIdError(
                    '\n  Problem finding worksheet by the name or id {0!r}\n'
                    '  within the Google spreadsheet {1!r}\n'
                    'Trying {0!r} as the id resulted in the following '
                    'RequestError:\n  {2}\n'
                    'Trying {0!r} as the name resulted in the following '
                    'SpreadSheetNameError:\n  {3}'.format(
                            worksheet_name_or_id,
                            self.spreadsheet.title.text,
                            id_error,
                            name_error))
            pass
        self._worksheet_id = wid 

    worksheet_id = property(_get_worksheet_id, _set_worksheet_id)

    def _get_spreadsheet(self):
        return self.spreadsheet_client.GetSpreadsheetsFeed(
                key = self.spreadsheet_key)

    spreadsheet = property(_get_spreadsheet)

    def _get_worksheet(self):
        return self.spreadsheet_client.GetWorksheetsFeed(
                key = self.spreadsheet_key,
                wksht_id = self.worksheet_id)

    worksheet = property(_get_worksheet)

    def _get_rows(self):
        row_feed = self.spreadsheet_client.GetListFeed(
                key = self.spreadsheet_key,
                wksht_id = self.worksheet_id)
        return row_feed.entry
    
    rows = property(_get_rows)

    def _get_cells(self):
        cell_feed = self.spreadsheet_key.GetCellsFeed(
                key = self.spreadsheet_key,
                wksht_id = self.worksheet_id)
        return cell_feed.entry

    cells = property(_get_cells)

    def _get_column_headers(self):
        return self.rows[0].custom.keys()

    column_headers = property(_get_column_headers)
        
    def _get_spreadsheet_key_by_name(self, spreadsheet_name):
        feed = self.spreadsheet_client.GetSpreadsheetsFeed()
        matches = []
        for spreadsheet in feed.entry:
            if spreadsheet.title.text == spreadsheet_name:
                matches.append(spreadsheet)
        if len(matches) < 1:
            raise SpreadsheetNameError(
                    'Could not find Google spreadsheet with name '
                    '{0!r}'.format(spreadsheet_name))
        elif len(matches) > 1:
            raise SpreadsheetNameError(
                    'Found multiple Google spreadsheets with name '
                    '{0!r}'.format(spreadsheet_name))
        return matches[0].id.text.split('/')[-1]

    def _get_worksheet_id_by_name(self, worksheet_name):
        feed = self.spreadsheet_client.GetWorksheetsFeed(
                key=self.spreadsheet_key)
        matches = []
        for worksheet in feed.entry:
            if worksheet.title.text == worksheet_name:
                matches.append(spreadsheet)
        if len(matches) < 1:
            raise WorksheetNameError(
                    'Could not find worksheet with name '
                    '{0!r} within Google spreadsheet '
                    '{1!r}'.format(
                            worksheet_name,
                            self.spreadsheet.title.text))
        elif len(matches) > 1:
            raise WorksheetNameError(
                    'Found multiple worksheets with name '
                    '{0!r} within Google spreadsheet '
                    '{1!r}'.format(
                            worksheet_name,
                            self.spreadsheet.title.text))
        return matches[0].id.text.split('/')[-1]

    def download_spreadsheet(self,
            path, 
            format='tsv'):
        entry = self.spreadsheet_client.GetSpreadsheetsFeed(
                key = self._spreadsheet_key)
        d_token = self.doc_client.GetClientLoginToken()
        self.doc_client.SetClientLoginToken(
                self.spreadsheet_client.GetClientLoginToken())
        self.doc_client.Download(
                entry_or_id_or_url=entry,
                file_path=path,
                export_format=format)
        self.doc_client.SetClientLoginToken(d_token)

    def _get_doc_entry_by_name(self, doc_name):
        q = gdata.docs.service.DocumentQuery()
        q['title'] = doc_name
        q['title-exact'] = 'true'
        feed = self.doc_client.Query(q.ToUri())
        if len(feed.entry) < 1:
            raise Exception('Could not find Google doc with name '
                    '{0!r}'.format(doc_name))
        elif len(feed.entry) > 1:
            raise Exception('Found multiple Google docs with name '
                    '{0!r}'.format(doc_name))
        return feed.entry[0]

    def googlize_column_header(self, header):
        return header.lower().replace(' ', '').replace('_', '')

    def filter_rows(self, column_headers, pattern):
        p = re.compile(pattern)
        headers = [self.googlize_column_header(x) for x in column_headers]
        for i, row in enumerate(self.rows):
            for h in headers:
                cell = row.custom.get(h)
                if cell.text and p.match(cell.text):
                    yield i, row

    def update_sheet(self, filter_headers, filter_pattern, update_header, update,
            only_update_empty_cells=False, dry_run=False):
        f_headers = [self.googlize_column_header(x) for x in filter_headers]
        u_header = self.googlize_column_header(update_header)
        for i, row in self.filter_rows(f_headers, filter_pattern):
            new_dict = {}
            skip = False
            for col_header, value in row.custom.iteritems():
                if col_header == u_header:
                    if only_update_empty_cells and value.text:
                        _LOG.warning('You requested only empty cells be '
                                'updated,\nbut row {0} matched the filter '
                                'criteria while having {1} in the target '
                                'cell\nSkipping!!'.format(i, value.text))
                        skip = True
                        break
                    new_dict[col_header] = update
                    _LOG.info('Row {0}. '
                              '{1}: {2} --> '
                              '{1}: {3}'.format(
                                  i, col_header, value.text, update))
                else:
                    new_dict[col_header] = value.text
            if not dry_run and not skip:
                self.update_row(row, new_dict) 

    def update_row(self, row, replacement_dict):
        self.spreadsheet_client.UpdateRow(row, replacement_dict)

    def insert_row(self, row_dict):
        return self.spreadsheet_client.InsertRow(
                row_dict,
                key = self.spreadsheet_key,
                wksht_id = self.worksheet_id)

