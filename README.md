latex-mailmerge
===============

This is a very simplistic tool that takes a Latex document with Python Code embedded between \begin{python}\end{python} and a CSV file and evaluates all the Python code for every line in the CSV file. During the evaluation of the embedded code, a variable for every column of the CSV file is available that contains the corresponding value at the current row.

For every row in the CSV file the part of the Latex document between \begin{document} and \end{document} is copied with the results of the Python evaluation pasted in.

The Python snippets don't share any state between them. There are no security measures whatsoever taken to prevent a malicious document from eating your data.

I use it to generate certificates for course attendance.