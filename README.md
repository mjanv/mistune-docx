# Mistune Docx Renderer

This tool **merges a list of Markdown files into a brand new Docx file !**

A template docx can be specified which implies that:
* The template defaults styles will be used and can be tweaked to change the appereance of the final document (This template needs to have the following styles defined: `BasicUserQuote` for blockquotes, `BasicUserList` for lists, `BasicUserTable` for tables)
* The template contents will be used as the first page(s) of the final documents 

Markdown supported features:
* Paragraphs
* Bold and italic text
* Headers (level 1 to 5)
* Bullet lists
* Tables with cells containing simple text (normal, bold or italic)
* Blockquote
* Images
* Mathematical equations (using Sympy rendering)
* Link
* Page break

Every feature follow the common Markdown syntax. A list of examples can be seen into the `example\` folder. Please remind that special characters such as `%` or `"` need to be espaced with a `\` to be admissible.

## How to use

```
python generate_doc.py output.docx --template example/template.docx --files example/*.md
```
The command requests:
* The name and location of the output file
* The name and location of the template file. If not specified, uses the defaults styles of the local Word installation and an empty document.
* The regex to locate the markdown files. If multiples Markdown files fits the specified regex, they will be prior assembled into one Markdown file following the alphabetic order.

## Requirements

This tool has been tested with Python 3.5. See _requirements.txt_ for all the libraries dependencies.
