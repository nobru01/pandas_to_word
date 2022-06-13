# import pandas as pd
# import docx


def tab_doc(df, doc_file_path):          # doc_file_path exemplo: '..\\output\\reprodutibilidade\\relatório de reprodutibilidade.docx'
    # open an existing document
    doc = docx.Document(doc_file_path)

    # add a table to the end and create a reference variable
    # extra row is so we can add the header row
    t = doc.add_table(df.shape[0]+1, df.shape[1])

    # add the header rows.
    for j in range(df.shape[-1]):
        t.cell(0,j).text = df.columns[j]

    # add the rest of the data frame
    for i in range(df.shape[0]):
        for j in range(df.shape[-1]):
            t.cell(i+1,j).text = str(df.values[i,j])

    # save the doc
    doc.save(doc_file_path)
    return 0