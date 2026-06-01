def replace_in_table(table, find, replace):
    for row in table.rows:
        for cell in row.cells:
            for subtable in cell.tables:
                if replace_in_table(subtable, find, replace):
                    return True
            for paragraph in cell.paragraphs:
                foundInRun = False
                for run in paragraph.runs:
                    if find in run.text:
                        if isinstance(replace, str):
                            run.text = run.text.replace(find, replace)
                            run.font.highlight_color = 0
                        else:
                            run.text = ""
                            paragraph._p.append(replace)
                        return True

                if not foundInRun and find in paragraph.text:
                    if isinstance(replace, str):
                        paragraph.text = paragraph.text.replace(find, replace)
                    else:
                        pass
                        paragraph._p.addnext(replace)

                    return True
    return False


if __name__ == "__main__":
    pass
