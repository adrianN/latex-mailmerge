# -*- coding: utf-8 -*-
import sys
import re
import csv
from collections import defaultdict

class OOCalc(csv.excel):
  delimiter = ';'

def parse(filename, d):
    f = open(filename,'r+b')
    reader = csv.DictReader(f, dialect = d)
    rows = list(reader)
    values = defaultdict(list)
    for row in rows:
        for (header, value) in  row.iteritems():
            values[header].append(value)
            
    f.close()
    return values


def eval_region(text, vars):
    try:
        return str(eval(text,vars))
    except:
        exec(text,vars)
        return str(vars["text"])
        


def fill_template(template, vars):
    blocks = [] #blocks of text
    pos = 0
    pattern = re.compile(r'\\begin\{python\}(.*?)\\end\{python\}',re.DOTALL)
    for match in pattern.finditer(template):
        blocks.append(template[pos:match.start()])
        pos = match.end()

        python = match.group(1)
        try:
            blocks.append(eval_region(python,vars))
        except:
            print "Can't successfully execute code block starting at", match.start(1)
            print python
            raise

    blocks.append(template[pos:])
    return "".join(blocks)

def all_copies(template, var_table):
    copies = []
    for i in xrange(min((len(variables[column]) for column in var_table))):
        v = {}
        for k in var_table.iterkeys():
            v[k] = var_table[k][i]

        copies.append(fill_template(template,v))
        
    return copies

def produce_tex(text, var_table):
    pattern = re.compile(r'\\begin\{document\}(.*?)\\end\{document\}', re.DOTALL)

    match = pattern.search(text)
    if not match:
       print "apparently not a valid latex file, no begin/end document found"
       exit
    preamble = text[0:match.start(1)]
    copies = all_copies(match.group(1), var_table)
    endamble = text[match.end(1):]
    return preamble + "\n\\newpage\n".join(copies) + endamble


if len(sys.argv) < 3:
    print "Usage: python mailmerge.py [-oocalc] template.tex data.csv"
    print "The -oocalc switch makes this work with csv produced by oocalc"
    sys.exit(-1)

try:
  sys.argv.remove('-oocalc')
  dialect = OOCalc
except ValueError:
  dialect = csv.excel
    
template_file = sys.argv[1]
variables_table = sys.argv[2]

print "Producing output from template file", template_file, "and data", variables_table

f = open(template_file, 'r+b')
template_text = f. read()
f.close()
variables = parse(variables_table, dialect)
print 'Available variables', variables.keys()

print "Will produce",str(min((len(variables[column]) for column in variables))),"pages output"

output = produce_tex(template_text, variables)
f = open('out.tex','wb')
f.write(output)
f.close()
