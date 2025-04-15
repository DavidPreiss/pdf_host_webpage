import os

pdf_path = './pdf_folder'

3 = [f for f in os.listdir(pdf_path) if f.endswith('.pdf')]

html_content ="<!DOCTYPE html>\n<html>\n<head>\n<title>PDF Files</title>\n</head>\n<body>\n<h2>All PDFs</h2>\n<ul>\n"
for pdf in pdf_files:
    html_content += f'    <li><a target="_blank" href="{pdf_path}/{pdf}">{pdf}</a></li>\n'
html_content += "</ul>\n</body>\n</html>"

pageName = "test.html"

with open(pageName, "w") as file:
    file.write(html_content)

print(f"{pageName} has been generated.")

