
def create_doc_latex(costs, totals,filename):
    latex_content = r"""
    \documentclass[10pt, a4paper]{article}
    \usepackage[a4paper,left=0.5cm,right=0.5cm,top=1cm,bottom=2cm]{geometry}
    \usepackage{array}
    \usepackage{booktabs}
    \usepackage{longtable}

    \begin{document}
    % Define column types with paragraph wrapping
    \newcolumntype{L}{>{\raggedright\arraybackslash}p{5cm}}  % Category column
    \newcolumntype{J}{>{\raggedright\arraybackslash}p{10cm}}  % Description and Justification column

    \section*{Budget Justification}

    % Begin the longtable environment
    \begin{longtable}{LJr}
    \toprule
    \textbf{Category} & \textbf{Description and Justification} & \hfill \textbf{Cost} \\
    \midrule
    \endhead
    """

    for year, items in costs.items():
        latex_content += "\\multicolumn{3}{c}{{\\textbf{{"+f"{year}"+"}}}} \\\\\n"
        latex_content += "\\midrule\n"
        for name, description, cost in items:
            if not int(cost):
                continue
            description = description.split("Unimelb: ")[0] if "Unimelb: " in description else description
            description = description.replace("$", "\$").replace('%', '\%').replace('&','& \small ')
            # Escape special LaTeX characters in the output
            name = name.replace('&', '\&').replace('%', '\%')
            latex_content += f"\small {name} & \small {description} & \small \${int(cost):,} \\\\\n"
        latex_content += f"\\textbf{{Total}} & & \\textbf{{\${int(totals[year]):,}}} \\\\\n"
        latex_content += "\\midrule\n"

    latex_content += "\\bottomrule\n"
    latex_content += "\\end{longtable}\n"
    latex_content += f"\\section*{{Grand Total: \${int(sum(totals.values())):,}}}\n"
    latex_content += r"\end{document}"

    #write to budget-arc.tex
    with open(filename, 'w') as file:
        file.write(latex_content)

    import subprocess
    subprocess.run(["lualatex", filename])


def create_doc_markdown(costs, totals, output_path):
    # Creating the markdown content
    markdown_content = "\n"
    for year, items in costs.items():
        #markdown_content += f"### {year} Budget Justification\n\n"
        markdown_content += "| Category            | Description and Justification                               | Cost        |\n"
        markdown_content += "|---------------------|-------------------------------------------|-------------|\n"
        for name, description, cost in items:
            markdown_content += f"| {name} | {description} | ${int(cost):,} |\n"
        markdown_content += f"| **Total** |                                                          | ${int(totals[year]):,} |\n\n"
        markdown_content += f": {year} Budget Justification\n\n"

    markdown_content += f"### Grand Total: ${int(sum(totals.values())):,}\n"

    print(markdown_content)

    # Write the markdown content to a file
    with open(output_path, 'w') as file:
        file.write(markdown_content)


def create_doc_pandoc(costs, totals):
    doc_elements = []
  # Create Markdown content dynamically based on provided data
    for year, items in costs.items():
        year_header = pf.Header(pf.Str(f"{year} Budget Justification"), level=2)
        doc_elements.append(year_header)

        # Table Header
        header_row = pf.TableRow(
            pf.TableCell(pf.Para(pf.Str("Category"))),
            pf.TableCell(pf.Para(pf.Str("Description"))),
            pf.TableCell(pf.Para(pf.Str("Cost")))
        )
        table_head = pf.TableHead(header_row)

        # Data Rows
        data_rows = []
        for name, description, cost in items:
            data_rows.append(
                pf.TableRow(
                    pf.TableCell(pf.Para(pf.Str(name))),
                    pf.TableCell(pf.Para(pf.Str(description))),
                    pf.TableCell(pf.Para(pf.Str(f"${int(cost):,}")))
                )
            )

        # Total Row
        table_foot = pf.TableFoot(pf.TableRow(
            pf.TableCell(pf.Para(pf.Strong(pf.Str("Total"))), colspan=2),
            pf.TableCell(pf.Para(pf.Str(f"${int(totals[year]):,}")))
        ))
        #data_rows.append(total_row)

        # Assemble Table Body
        table_body = pf.TableBody(*data_rows)

        # Assemble Table
        table = pf.Table(table_body, head=table_head, foot=table_foot)
        doc_elements.append(table)

    # Grand Total
    grand_total = pf.Para(pf.Strong(pf.Str(f"Grand Total: ${int(sum(totals.values())):,}")))
    doc_elements.append(grand_total)

    # Create the final document
    doc = pf.Doc(*doc_elements)

    with open('output.json', 'w', encoding='utf-8') as f:
        pf.dump(doc, f)

    # run pandoc in a subprocess to produce budget-arc.pdf
    import subprocess
    subprocess.run(["pandoc", "-V", "geometry:margin=1in", "-V", "fontsize=10pt", "output.json", "-o", "budget-arc.docx"])
