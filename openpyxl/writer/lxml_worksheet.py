from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

# Experimental writer of worksheet data using lxml incremental API

from io import BytesIO
from lxml.etree import xmlfile, Element, SubElement, fromstring

from openpyxl.compat import (
    itervalues,
    safe_string,
    iteritems
)
from openpyxl.cell import (
    column_index_from_string,
    coordinate_from_string,
)
from openpyxl.xml.constants import (
    REL_NS,
    SHEET_MAIN_NS
)

from openpyxl.formatting import ConditionalFormatting
from openpyxl.worksheet.datavalidation import writer

from .worksheet import (
    row_sort,
    get_rows_to_write,
    write_datavalidation,
)


def write_worksheet(worksheet, shared_strings):
    """Write a worksheet to an xml file."""
    vba_attrs = {}
    vba_root = None
    if worksheet.xml_source:
        vba_root = fromstring(worksheet.xml_source)
        el = vba_root.find('{%s}sheetPr' % SHEET_MAIN_NS)
        if el is not None:
            vba_attrs['codeName'] = el.get('codeName', worksheet.title)

    out = BytesIO()
    NSMAP = {None : SHEET_MAIN_NS}
    with xmlfile(out) as xf:
        with xf.element('worksheet', nsmap=NSMAP):

            pr = Element('sheetPr', vba_attrs)
            SubElement(pr, 'outlinePr',
                       {'summaryBelow':
                        '%d' %  (worksheet.show_summary_below),
                        'summaryRight': '%d' %     (worksheet.show_summary_right)})
            if worksheet.page_setup.fitToPage:
                SubElement(pr, 'pageSetUpPr', {'fitToPage': '1'})
            xf.write(pr)

            dim = Element('dimension', {'ref': '%s' % worksheet.calculate_dimension()})
            xf.write(dim)

            xf.write(write_sheetviews(worksheet))
            xf.write(write_format(worksheet))
            cols = write_cols(worksheet)
            if cols is not None:
                xf.write(cols)
            write_rows(xf, worksheet)

            if worksheet.protection.sheet:
                prot = Element('sheetProtection', dict(worksheet.protection))
                xf.write(prot)
                del prot

            af = write_autofilter(worksheet)
            if af is not None:
                xf.write(af)

            merge = write_mergecells(worksheet)
            if merge is not None:
                xf.write(merge)

            write_conditional_formatting(xf, worksheet)
            dv = write_datavalidation(worksheet)
            if dv is not None:
                xf.write(dv)

            hyper = write_hyperlinks(worksheet)
            if hyper is not None:
                xf.write(hyper)

            options = worksheet.page_setup.options
            if options:
                print_options = Element('printOptions', options)
                xf.write(print_options)
                del print_options

            margins = Element('pageMargins', dict(worksheet.page_margins))
            xf.write(margins)
            del margins

            setup = worksheet.page_setup.setup
            if setup:
                page_setup = Element('pageSetup', setup)
                xf.write(page_setup)
                del page_setup

            hf = write_header_footer(worksheet)
            if hf is not None:
                xf.write(hf)

            if worksheet._charts or worksheet._images:
                drawing = Element('drawing', {'{%s}id' % REL_NS: 'rId1'})
                xf.write(drawing)
                del drawing

            # If vba is being preserved then add a legacyDrawing element so
            # that any controls can be drawn.
            if vba_root is not None:
                el = vba_root.find('{%s}legacyDrawing' % SHEET_MAIN_NS)
                if el is not None:
                    rId = el.get('{%s}id' % REL_NS)
                    legacy = Element('legacyDrawing', {'{%s}id' % REL_NS: rId})
                    xf.write(legacy)

            pb = write_pagebreaks(worksheet)
            if pb is not None:
                xf.write(pb)

            # add a legacyDrawing so that excel can draw comments
            if worksheet._comment_count > 0:
                comments = Element('legacyDrawing', {'{%s}id' % REL_NS: 'commentsvml'})
                xf.write(comments)

    xml = out.getvalue()
    out.close()
    return xml


def write_cols(worksheet):
    """Write worksheet columns to xml.

    <cols> may never be empty -
    spec says must contain at least one child
    """
    cols = []
    for label, dimension in iteritems(worksheet.column_dimensions):
        dimension.style = worksheet._styles.get(label)
        col_def = dict(dimension)
        if col_def == {}:
            continue
        idx = column_index_from_string(label)
        cols.append((idx, col_def))

    if not cols:
        return

    el = Element('cols')

    for idx, col_def in sorted(cols):
        v = "%d" % idx
        cmin = col_def.get('min') or v
        cmax = col_def.get('max') or v
        col_def.update({'min': cmin, 'max': cmax})
        el.append(Element('col', col_def))
    return el


def write_rows(xf, worksheet):
    """Write worksheet data to xml."""

    cells_by_row = get_rows_to_write(worksheet)

    with xf.element("sheetData"):
        for row_idx in sorted(cells_by_row):
            # row meta data
            row_dimension = worksheet.row_dimensions[row_idx]
            row_dimension.style = worksheet._styles.get(row_idx)
            attrs = {'r': '%d' % row_idx,
                     'spans': '1:%d' % worksheet.max_column}
            attrs.update(dict(row_dimension))

            with xf.element("row", attrs):

                row_cells = cells_by_row[row_idx]
                for cell in sorted(row_cells, key=row_sort):
                    write_cell(xf, worksheet, cell)


def write_cell(xf, worksheet, cell):
    string_table = worksheet.parent.shared_strings
    coordinate = cell.coordinate
    attributes = {'r': coordinate}
    if cell.has_style:
        attributes['s'] = '%d' % cell._style

    if cell.data_type != 'f':
        attributes['t'] = cell.data_type

    value = cell.internal_value

    if value in ('', None):
        with xf.element("c", attributes):
            return

    with xf.element('c', attributes):
        if cell.data_type == 'f':
            shared_formula = worksheet.formula_attributes.get(coordinate, {})
            if shared_formula is not None:
                if (shared_formula.get('t') == 'shared'
                    and 'ref' not in shared_formula):
                    value = None
            with xf.element('f', shared_formula):
                if value is not None:
                    xf.write(value[1:])
                    value = None

        if cell.data_type == 's':
            value = string_table.add(value)
        with xf.element("v"):
            if value is not None:
                xf.write(safe_string(value))


def write_autofilter(worksheet):
    auto_filter = worksheet.auto_filter
    if auto_filter.ref is None:
        return

    el = Element('autoFilter', {'ref': auto_filter.ref})
    if (auto_filter.filter_columns
        or auto_filter.sort_conditions):
        for col_id, filter_column in sorted(auto_filter.filter_columns.items()):
            fc = SubElement(el, 'filterColumn', {'colId': str(col_id)})
            attrs = {}
            if filter_column.blank:
                attrs = {'blank': '1'}
            flt = SubElement(fc, 'filters', attrs)
            for val in filter_column.vals:
                SubElement(flt, 'filter', {'val': val})
        if auto_filter.sort_conditions:
            srt = SubElement(el,  'sortState', {'ref': auto_filter.ref})
            for sort_condition in auto_filter.sort_conditions:
                sort_attr = {'ref': sort_condition.ref}
                if sort_condition.descending:
                    sort_attr['descending'] = '1'
                SubElement(srt, 'sortCondtion', sort_attr)
    return el


def write_sheetviews(worksheet):
    views = Element('sheetViews')
    sheetviewAttrs = {'workbookViewId': '0'}
    if not worksheet.show_gridlines:
        sheetviewAttrs['showGridLines'] = '0'
    view = SubElement(views, 'sheetView', sheetviewAttrs)
    selectionAttrs = {}
    topLeftCell = worksheet.freeze_panes
    if topLeftCell:
        colName, row = coordinate_from_string(topLeftCell)
        column = column_index_from_string(colName)
        pane = 'topRight'
        paneAttrs = {}
        if column > 1:
            paneAttrs['xSplit'] = str(column - 1)
        if row > 1:
            paneAttrs['ySplit'] = str(row - 1)
            pane = 'bottomLeft'
            if column > 1:
                pane = 'bottomRight'
        paneAttrs.update(dict(topLeftCell=topLeftCell,
                              activePane=pane,
                              state='frozen'))
        SubElement(view, 'pane', paneAttrs)
        selectionAttrs['pane'] = pane
        if row > 1 and column > 1:
            SubElement(view, 'selection', {'pane': 'topRight'})
            SubElement(view, 'selection', {'pane': 'bottomLeft'})

    selectionAttrs.update({'activeCell': worksheet.active_cell,
                           'sqref': worksheet.selected_cell})

    SubElement(view, 'selection', selectionAttrs)
    return views


def write_format(worksheet):
    attrs = {'defaultRowHeight': '15', 'baseColWidth': '10'}
    dimensions_outline = [dim.outline_level
                          for dim in itervalues(worksheet.column_dimensions)]
    if dimensions_outline:
        outline_level = max(dimensions_outline)
        if outline_level:
            attrs['outlineLevelCol'] = str(outline_level)
    return Element('sheetFormatPr', attrs)


def write_mergecells(worksheet):
    """Write merged cells to xml."""
    cells = worksheet._merged_cells
    if not cells:
        return

    merge = Element('mergeCells', {'count':'%d' % len(cells)})
    for range_string in cells:
        attrs = {'ref': range_string}
        SubElement(merge, 'mergeCell', attrs)
    return merge


def write_header_footer(worksheet):
    header = worksheet.header_footer.getHeader()
    footer = worksheet.header_footer.getFooter()
    if header or footer:
        tag = Element('headerFooter')
        if header:
            SubElement(tag, 'oddHeader').text = header
        if worksheet.header_footer.hasFooter():
            SubElement(tag, 'oddFooter').text = footer
        return tag


def write_pagebreaks(worksheet):
    breaks = worksheet.page_breaks
    if breaks:
        tag = Element( 'rowBreaks', {'count': str(len(breaks)),
                                     'manualBreakCount': str(len(breaks))})
        for b in breaks:
            tag.append(Element('brk', {'id': str(b), 'man': 'true', 'max': '16383',
                             'min': '0'}))
        return tag


def write_hyperlinks(worksheet):
    """Write worksheet hyperlinks to xml."""
    tag = Element('hyperlinks')
    for cell in worksheet.get_cell_collection():
        if cell.hyperlink_rel_id is not None:
            attrs = {'display': cell.hyperlink,
                     'ref': cell.coordinate,
                     '{%s}id' % REL_NS: cell.hyperlink_rel_id}
            SubElement(tag, 'hyperlink', attrs)
    if tag.getchildren():
        return tag


def write_conditional_formatting(xf, worksheet):
    """Write conditional formatting to xml."""
    for range_string, rules in iteritems(worksheet.conditional_formatting.cf_rules):
        if not len(rules):
            # Skip if there are no rules.  This is possible if a dataBar rule was read in and ignored.
            continue
        cf = Element('conditionalFormatting', {'sqref': range_string})
        for rule in rules:
            if rule['type'] == 'dataBar':
                # Ignore - uses extLst tag which is currently unsupported.
                continue
            attr = {'type': rule['type']}
            for rule_attr in ConditionalFormatting.rule_attributes:
                if rule_attr in rule:
                    attr[rule_attr] = str(rule[rule_attr])
            cfr = SubElement(cf, 'cfRule', attr)
            if 'formula' in rule:
                for f in rule['formula']:
                    SubElement(cfr, 'formula').text = f
            if 'colorScale' in rule:
                cs = SubElement(cfr, 'colorScale')
                for cfvo in rule['colorScale']['cfvo']:
                    SubElement(cs, 'cfvo', cfvo)
                for color in rule['colorScale']['color']:
                    SubElement(cs, 'color', dict(color))
            if 'iconSet' in rule:
                iconAttr = {}
                for icon_attr in ConditionalFormatting.icon_attributes:
                    if icon_attr in rule['iconSet']:
                        iconAttr[icon_attr] = rule['iconSet'][icon_attr]
                iconSet = SubElement(cfr, 'iconSet', iconAttr)
                for cfvo in rule['iconSet']['cfvo']:
                    SubElement(iconSet, 'cfvo', cfvo)
        xf.write(cf)