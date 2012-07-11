# -*- coding: utf-8 -*-
import sys
import re
import csv
from collections import defaultdict
import string
from optparse import OptionParser

class OOCalc(csv.excel):
  delimiter = ';'

def parse(filename, d):
    f = open(filename,'rb')
    reader = csv.DictReader(f, dialect = d)
    rows = list(reader)
    values = defaultdict(list)
    for row in rows:
        for (header, value) in  row.iteritems():
            values[header].append(value)
            
    f.close()
    return values

def strip_evil_whitespace(text):
  def num_leading_whitespace(line):
    i=0
    while i<len(line) and line[i] in string.whitespace:
      i+=1
    return i

  lines = [l for l in text.splitlines() if len(l)>0]
  m = min(map(num_leading_whitespace, lines))
  for i in xrange(len(lines)):
    lines[i] = lines[i][m:]
  
  return '\n'.join(lines)


def eval_region(text, vars):
    text = strip_evil_whitespace(text)
    if options.dry:
      return r'\texttt{<Python>}'
    try:
        return str(eval(text,vars))
    except:
        exec(text,vars)
        return str(vars["_text"])
        


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
        v['_line'] = i
        try:
	    copies.append(fill_template(template,v))
        except:
            print "Error while processing line", i, "of the csv"
            raise
        
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


usage = "usage: python %prog [options] template.tex [data.csv]"
parser = OptionParser(usage)
parser.add_option('--oocalc',
                  action = 'store_true',
                  dest = 'oocalc',
                  default = False,
                  help = 'Switch csv format from ","-separated (Excel default) to ";"-separated (OOCalc default)')
parser.add_option('-d','--dry',
                  action='store_true',
                  dest = 'dry',
                  default = False,
                  help = 'Switch to dry run mode. Don\'t require a csv, replace code block by dummy strings')
parser.add_option('-o','--out',
                  dest = "out",
                  default = 'out.tex',
                  help = 'Set the output file. Default: %default') 



options, args = parser.parse_args(sys.argv)
if len(args) < 3 and not options.dry or len(args)<2:
    parser.print_help()
    sys.exit(-1)

if options.oocalc:
    options.dialect = OOCalc
else:
    options.dialect = csv.excel

if options.dry:
    print "Dry run mode"
    variables = {'dry':['dry']}
else:
    variables_table = args[2]
    variables = parse(variables_table, options.dialect)
    print 'Available variables', variables.keys()
 
template_file = args[1]
f = open(template_file, 'rb')
template_text = f. read()
f.close()

print "Producing output from template file", template_file,
if options.dry:
  print 
else:
  print "and data", variables_table
print "Storing things into", options.out
print "Will produce",str(min((len(variables[column]) for column in variables))),"\"pages\"(s) output"


output = produce_tex(template_text, variables)
f = open(options.out,'wb')
f.write(output)
f.close()
