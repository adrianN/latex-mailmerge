latex-mailmerge
===============

This is a very simplistic tool that takes a Latex document with Python Code embedded between \begin{python}\end{python} and a CSV file and evaluates all the Python code for every line in the CSV file. During the evaluation of the embedded code, a variable for every column of the CSV file is available that contains the corresponding value at the current row.

For every row in the CSV file the part of the Latex document between \begin{document} and \end{document} is copied with the results of the Python evaluation pasted in.

The Python snippets don't share any state between them. There are no security measures whatsoever taken to prevent a malicious document from eating your data.

I use it to generate certificates for course attendance.

Usage
-----

1. Prepare your CSV. There should be no empty lines and every column should have a header. If it's not Comma-separated, but Semikolon-separated (as produced by OOCalc) add the `-oocalc` switch at step 3.
2. Prepare your Latex template. It should be a normal Latex document, except you can have Python code between \begin{python} and \end{python}. In this code there are magic variables available: The headers of your columns and `text`. Python expressions can ignore `text`, Python statements should assign to `text`. See the `example` folder for an example template
3. Run `python mailmerge.py template.tex data.csv`. It produces `out.tex`. Try running without arguments to get some usage information.
4. Run (pdf)Latex on `out.tex`