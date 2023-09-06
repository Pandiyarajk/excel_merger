import xlrd
import xlwt
from xlutils.copy import copy
import os
import glob  # Filter files


def merge_files(input_file_1, input_file_2, output_file_name):
    # Load the first Excel workbook
    workbook1 = xlrd.open_workbook(input_file_1)

    # Load the second Excel workbook
    workbook2 = xlrd.open_workbook(input_file_2)

    # Create a new workbook to hold the merged sheets
    merged_workbook = xlwt.Workbook()

    for sheet_name in workbook1.sheet_names():
        if sheet_name in workbook2.sheet_names():
            # print("Merging sheet {}".format(sheet_name))
            sheet1 = workbook1.sheet_by_name(sheet_name)
            sheet2 = workbook2.sheet_by_name(sheet_name)
            merged_sheet = merged_workbook.add_sheet(sheet_name)

            # Copy content from the first file sheet
            for row_index in range(sheet1.nrows):
                for col_index in range(sheet1.ncols):
                    merged_sheet.write(row_index, col_index, sheet1.cell_value(row_index, col_index))

            # Copy content from the second file sheet starting from next row
            for row_index in range(1, sheet2.nrows):
                for col_index in range(sheet2.ncols):
                    merged_sheet.write(row_index + sheet1.nrows-1, col_index, sheet2.cell_value(row_index, col_index))
        else:
            # print("Sheet {} available only in file {}".format(sheet_name, input_file_1))
            sheet1 = workbook1.sheet_by_name(sheet_name)
            merged_sheet = merged_workbook.add_sheet(sheet_name)
            # Copy content from the first sheet
            for row_index in range(sheet1.nrows):
                for col_index in range(sheet1.ncols):
                    merged_sheet.write(row_index, col_index, sheet1.cell_value(row_index, col_index))

    for sheet_name in workbook2.sheet_names():
        if sheet_name not in workbook1.sheet_names():
            # print("Sheet {} available only in file {}".format(sheet_name, input_file_2))
            sheet2 = workbook2.sheet_by_name(sheet_name)
            merged_sheet = merged_workbook.add_sheet(sheet_name)
            # Copy content from the second file sheet
            for row_index in range(sheet2.nrows):
                for col_index in range(sheet2.ncols):
                    merged_sheet.write(row_index, col_index, sheet2.cell_value(row_index, col_index))

    # Save the merged workbook
    merged_workbook.save(output_file_name)

    print("Sheets merged successfully.")


def sort_rows_by_first_column(input_file_name, output_file_name):
    # Load the Excel workbook
    workbook = xlrd.open_workbook(input_file_name, formatting_info=True)

    # Create a copy of the workbook to write changes
    workbook_copy = copy(workbook)

    for sheet_index, sheet_name in enumerate(workbook.sheet_names()):
        sheet = workbook.sheet_by_index(sheet_index)
        sorted_indices = sorted(range(1, sheet.nrows), key=lambda x: sheet.cell_value(x, 0))
        sorted_indices.insert(0, 0)  # Include the header row

        sorted_sheet = workbook_copy.get_sheet(sheet_index)

        for new_row_index, old_row_index in enumerate(sorted_indices):
            for col_index in range(sheet.ncols):
                sorted_sheet.write(new_row_index, col_index, sheet.cell_value(old_row_index, col_index))

    # Save the changes
    workbook_copy.save(output_file_name)

    print("Rows sorted by first column value in all sheets.")


def remove_duplicate_rows_by_first_column(input_file_name, output_file_name):
    # Load the Excel workbook
    workbook = xlrd.open_workbook(input_file_name, formatting_info=True)

    # Create a copy of the workbook to write changes
    workbook_copy = copy(workbook)

    for sheet_index, sheet_name in enumerate(workbook.sheet_names()):
        sheet = workbook.sheet_by_index(sheet_index)
        seen_values = set()
        rows_to_keep = []

        for row_index, row in enumerate(sheet.get_rows()):
            if row_index == 0:  # Skip the header row
                continue
            value = row[0].value
            if value in seen_values:
                ...
            else:
                rows_to_keep.append(row_index)
                seen_values.add(value)

        new_sheet = workbook_copy.get_sheet(sheet_index)

        # Skip header by setting start index by 1
        new_row_index = 1
        # Move unique rows to top
        for rtk in rows_to_keep:
            for col_index in range(sheet.ncols):
                new_sheet.write(new_row_index, col_index, sheet.cell_value(rtk, col_index))
            new_row_index = new_row_index + 1
        # Empty the remaining cells
        for empty_index in range(new_row_index, sheet.nrows):
            for col_index in range(sheet.ncols):
                new_sheet.write(empty_index, col_index, "")

        if sheet.nrows-new_row_index > 0:
            print("{}: cleared {} rows.".format(sheet_name, sheet.nrows-new_row_index))

    # Save the changes
    workbook_copy.save(output_file_name)

    print("Duplicate rows removed by first column value in all sheets.")


if __name__ == "__main__":
    # Merge sheets of all files in the directory
    merged_file = ""
    xls_files = glob.glob(os.path.join(os.getcwd(), "*.xls"))
    file_list = [os.path.basename(full_name) for full_name in xls_files]

    for index, file in enumerate(file_list):
        merged_file = "merged_file" + str(index) + ".xls"
        if index == 1:
            print("Merging files {} and {}".format(file_list[index-1], file_list[index]))
            merge_files(file_list[index-1], file_list[index], merged_file)
        elif index > 1:
            print("Merging file {}".format(file_list[index]))
            prev_merged_file = "merged_file" + str(index-1) + ".xls"
            merge_files(prev_merged_file, file_list[index], merged_file)
            os.remove(prev_merged_file)
    print(merged_file)

    print("\n*** Sorting rows on all sheets ***")
    # Sort rows
    sorted_file_name = 'sorted_file.xls'
    sort_rows_by_first_column(merged_file, sorted_file_name)
    os.remove(merged_file)

    print("\n*** Removing duplicate rows on all sheets ***")
    # Remove duplicates
    clean_file_name = 'merged_file.xls'
    remove_duplicate_rows_by_first_column(sorted_file_name, clean_file_name)
    os.remove(sorted_file_name)
