from openpyxl import Workbook, load_workbook
from openpyxl.utils import rows_from_range, get_column_letter
from config import HEADERS, ROW_START, ROW_ITEM_COUNT, ROW_GAP, VAT_POSITIONS


# TODO: Cache the result of merged_cells
# Returns the value for a cell. If multiple cells are merged,
# the merge range will have the same value
def get_value_with_merge_lookup(sheet, cell, merged_cells_ranges):
    # EG: A12, B12
    idx = cell.coordinate

    for range_ in merged_cells_ranges:
        merged_cells = list(rows_from_range(str(range_)))
        for row in merged_cells:
            if idx in row:
                return str(sheet[merged_cells[0][0]].value)
    return str(sheet[idx].value)


# Returns the value that will be added in the new worksheet
# Based on the values extracted from the original
def get_cell_value_for_new_worksheet(header_item, header, values, vat):
    value = None
    if header_item['value'] is not None:
        value = header_item['value']
    else:
        if header_item['position_original_document'] is not None:
            value = values[header_item['position_original_document']]
        # Special scenario for VAT specific fields
        else:
            if header == 'Valoare neta totala':
                value = vat['total_net_amount']
            elif header == 'Valoare TVA':
                value = vat['vat']
            elif header == 'Total document':
                value = vat['total_amount']
    return value


# Do not consider invoices with a total ammount of 0
def invoice_should_be_processed(values):
    return float(values[12]) != 0


# Get the VATs that are in the curent row
def get_vat_rates(values):
    return [{
        'total_net_amount':
        values[vat['total_amount']],
        'vat':
        values[vat['vat']],
        'total_amount':
        "{:.2f}".format(
            float(values[vat['total_amount']]) + float(values[vat['vat']]))
    } for vat in VAT_POSITIONS.values()
            if float(values[vat['total_amount']]) > 0.00]


# Original document
wb = load_workbook('raport.xlsx')
ws = wb.active
merged_cells_ranges = ws.merged_cells.ranges

# New document
new_wb = Workbook()
new_ws = new_wb.active

# Start processing
ok = True
row_new_document = 1
while ok == True:
    # Iterate over rows in the original document
    #TODO: Check needed for pages with less invoices ( eg: 20 )
    for row_original_document in range(ROW_START,
                                       ROW_START + ROW_ITEM_COUNT + 1):
        values = []
        # Going through each cell of a row in the original document
        # Get cell values from the original document after adjusting the merged cells
        for cell in ws[row_original_document]:
            values.append(
                get_value_with_merge_lookup(ws, cell, merged_cells_ranges))

        # If total ammount is 0, skip this row
        if not invoice_should_be_processed(values):
            continue

        # Detect the different rates for VAT
        vats = get_vat_rates(values)

        for vat in vats:
            # Start building the new file based on the specified headers
            for col, header in enumerate(HEADERS):
                header_item = HEADERS[header]
                value = get_cell_value_for_new_worksheet(
                    header_item, header, values, vat)
                char_new_ws = get_column_letter(col + 1)
                new_ws[char_new_ws + str(row_new_document)] = value
            # Process a new row in the new document
            row_new_document += 1
        ok = False

new_wb.save('newdocument.xlsx')
wb.close()
