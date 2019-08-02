from docx import Document
import pandas as pd
import argparse
import os


def get_table_fr_word(path):
    # Check the path to make sure docx is openned
    if path[-4:] == 'docx':
        # Open the docx document
        document = Document(path)
        table = document.tables[0]
        data = []
        # Get value in cells of table
        for i, row in enumerate(table.rows):
            text = [cell.text for cell in row.cells]
            # Get the list of column
            if i == 0:
                columns = text
                continue
            # Get the list of list of data (row)
            elif i > 0:
                data.append(text)
        # Create a table by using Panda lib
        df = pd.DataFrame(data=data,
                          columns=columns)
        cols = df.columns.tolist()  # get the list of columns
        # In case of table has more than 3 columns
        if len(cols) > 3:
            # Reorder the list by moving the column (index=1) to the front
            # and moving the column (index=2) to the end of table
            cols = [cols[1]] \
                    + list([a for a in cols
                            if a != cols[1] and a != cols[2]]) \
                    + [cols[2]]
            df = df[cols]  # Reorder the list of table with new list of column
        # In case of table has only two column, move the last to the first
        elif len(cols) == 2:
            cols = cols[-1:] + cols[:-1]
            df = df[cols]
        # In case of table has only one column, then raise Exception
        else:
            raise Exception('Your table only has one column')
        # Export table
        df.to_excel('output_table2.xls')
        return 'The table has been extracted successfully!\n'\
               'The output Excel file is stored at this path ~> {}'.format(os.getcwd())  # noqa
    else:
        return 'Can not open the file because the extension is not valid!\n'\
               'Verify that your path is going to open a docx file.'


def main():
    parser = argparse.ArgumentParser(
        description='A simple function to extract table from a DOCX file')
    parser.add_argument('path', help='path to a docx file', type=str)
    arg = parser.parse_args().path
    print(get_table_fr_word(arg))


if __name__ == "__main__":
    main()
