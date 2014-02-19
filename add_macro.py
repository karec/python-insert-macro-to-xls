__author__ = 'karec'
import zipfile
import os
import shutil
import tempfile
import sys
from xml.dom import minidom


def add_macro(fname, *filenames, **kwargs):
    """
    This function remove the filenames from archive
    and store them in dict for manipulation

    Args:
        fname (string): file to update with macro

    Filenames:
        name (string): string represent file to update or exclude from update
    
    Kwargs:
        macro_loc (string): absolute path of macro you want to insert
    
    """
    tempdir = tempfile.mkdtemp()
    macro_loc = 'vbaProject.bin'
    if kwargs.has_key('macro_loc') : macro_loc = kwargs['macro_loc']
    try:
        data = dict()
        tempname = os.path.join(tempdir, 'new.zip')
        with zipfile.ZipFile(fname, 'r') as zipread:
            with zipfile.ZipFile(tempname, 'w') as zipwrite:
                for item in zipread.infolist():
                    if item.filename not in filenames:
                        data = zipread.read(item.filename)
                        zipwrite.writestr(item, data)
                    else:
                        zipwrite.writestr(item, update_files(
                                item.filename, zipread.read(item.filename)
                                ))
                zipwrite.writestr('xl/vbaProject.bin', file(macro_loc).read())
        shutil.move(tempname, fname)
    finally:
        shutil.rmtree(tempdir)


def update_files(filename, data):
    """
    This function dispatch the data for return good value for each file
    """
    if '[Content_Types].xml' in filename:
        return update_content_types(data)
    elif 'workbook' in filename:
        return update_workbook(data)

def update_workbook(data):
    """
    This function append new tag for workbook
    """
    xml = minidom.parseString(data)
    rel = xml.getElementsByTagName('Relationships')[0]
    elem = xml.createElement('Relationship')
    elem.attributes['Id'] = 'rId99'
    elem.attributes['Type'] = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject'
    elem.attributes['Target'] = 'vbaProject.bin'
    rel.appendChild(elem)
    return xml.toxml()

def update_content_types(data):
    """
    This function append new macro relation in contenttypes file
    """
    xml = minidom.parseString(data)
    types = xml.getElementsByTagName('Types')[0]
    elem = xml.createElement('Override')
    elem.attributes['PartName'] = '/xl/vbaProject.bin'
    elem.attributes['ContentType'] = 'application/vnd.ms-office.vbaProject'
    types.appendChild(elem)
    return xml.toxml()

if __name__ == '__main__':
    f = raw_input('File name and location : ')
    if f == '': f = 'final.xlsx'
    add_macro(f, '[Content_Types].xml', 'xl/_rels/workbook.xml.rels')
