from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


import os.path

from openpyxl.comments import Comment, default_height, default_width
from openpyxl.xml.constants import (
    PACKAGE_WORKSHEET_RELS,
    SHEET_MAIN_NS,
    COMMENTS_NS,
    PACKAGE_XL,
    VML_NS,
    )
from openpyxl.writer.comments import excelns, officens, vmlns
from openpyxl.xml.functions import fromstring, safe_iterator
import string


def _get_author_list(root):
    author_subtree = root.find('{%s}authors' % SHEET_MAIN_NS)
    return [author.text for author in author_subtree]

def read_comments(ws, xml_source, drawing_xml_source=None):
    """Given a worksheet and the XML of its comments file, assigns comments to cells"""
    root = fromstring(xml_source)
    authors = _get_author_list(root)
    comment_nodes = list(safe_iterator(root, ('{%s}comment' % SHEET_MAIN_NS)))
    # pull all refs to create hash
    refs = [node.attrib['ref'] for node in comment_nodes]
    if drawing_xml_source:
        drawing_root = fromstring(drawing_xml_source)
        drawing_info = comments_drawing_info(refs, drawing_root)
    for node in comment_nodes:
        author = authors[int(node.attrib['authorId'])]
        cell = node.attrib['ref']
        height = default_height
        width = default_width
        # look up existing comment height/width
        if drawing_info:
            # is there drawing info specifically for this comment?
            comment_drawing_info = drawing_info.get(cell)
            if comment_drawing_info:
                height, width = comment_drawing_info
        text_node = node.find('{%s}text' % SHEET_MAIN_NS)
        substrs = []
        for run in text_node.findall('{%s}r' % SHEET_MAIN_NS):
            runtext = ''.join([t.text for t in run.findall('{%s}t' % SHEET_MAIN_NS)])
            substrs.append(runtext)
        comment_text = ''.join(substrs)
        comment = Comment(comment_text, author, default_height, default_width)
        ws.cell(coordinate=cell).comment = comment

def get_comments_file(worksheet_path, archive, valid_files):
    """Returns the XML filename in the archive which contains the comments for
    the spreadsheet with codename sheet_codename. Returns None if there is no
    such file"""
    sheet_codename = os.path.split(worksheet_path)[-1]
    rels_file = PACKAGE_WORKSHEET_RELS + '/' + sheet_codename + '.rels'
    if rels_file not in valid_files:
        return None
    rels_source = archive.read(rels_file)
    root = fromstring(rels_source)
    for i in root:
        if i.attrib['Type'] == COMMENTS_NS:
            comments_file = os.path.split(i.attrib['Target'])[-1]
            comments_file = PACKAGE_XL + '/' + comments_file
            if comments_file in valid_files:
                return comments_file
    return None


def get_drawings_file(worksheet_path, archive, valid_files):
    """Returns the XML filename in the archive which contains the drawings for
    the spreadsheet with codename sheet_codename. Returns None if there is no
    such file"""
    sheet_codename = os.path.split(worksheet_path)[-1]
    rels_file = PACKAGE_WORKSHEET_RELS + '/' + sheet_codename + '.rels'
    if rels_file not in valid_files:
        return None
    rels_source = archive.read(rels_file)
    root = fromstring(rels_source)
    for i in root:
        if i.attrib['Type'] == VML_NS:
            drawings_file = os.path.split(i.attrib['Target'])[-1]
            drawings_file = PACKAGE_XL + '/drawings/' + drawings_file
            if drawings_file in valid_files:
                return drawings_file
    return None

def comments_drawing_info(refs, xml_nodes_to_scan):
    """Returns drawing info for reference cells in drawing vml"""

    drawing_info = {}
    for ref in refs:
        alpha_column = ref[:1]
        column = string.ascii_lowercase.index(alpha_column.lower())
        height_str = default_height
        width_str = default_width
        shape_nodes = xml_nodes_to_scan.findall('{%s}shape' % vmlns)
        for node in shape_nodes:
            # get height and width from the style attribute
            client_data_node = node.find('{%s}ClientData' % excelns)
            column_node = client_data_node.find('{%s}Column' % excelns)
            # cast as int so it matches the index above
            excel_column = int(column_node.text)
            if excel_column == column:
                style_string = node.attrib['style'].replace(' ', '')
                height_width = height_width_from_style(style_string)
                drawing_info[ref] = height_width
    return drawing_info


def height_width_from_style(style_string):
    """Helper method to return height and width tuple from style string"""
    height_string = 'height:'
    width_string = 'width:'
    height_locator = style_string.find(height_string) + len(height_string)
    right_of_height = style_string[height_locator:]
    height_semic_locator = right_of_height.find(';')
    height_str = right_of_height[:height_semic_locator]

    width_locator = style_string.find(width_string) + len(width_string)
    right_of_width = style_string[width_locator:]
    width_semic_locator = right_of_width.find(';')
    width_str = right_of_width[:width_semic_locator]

    return (height_str, width_str)
