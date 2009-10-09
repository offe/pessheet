# Todo:
 # GetDotFile (GraphVis-file)
 # Labels
 # Actions
 # Some way of finding out when cells needs to be updated (listeners)?

import compiler
import re

class SpreadSheetError:
    def __init__(self, msg):
        self.msg = msg
    def __str__(self):
        return 'Error: %s' % self.msg 

class SpreadSheetCell:

    def __init__(self, name, formula, evaluator, supporting_cell=False):
        self._name = name
        self._updated = False
        self._evaluator = evaluator
        self._cashedvalue = None
        self._precedents = set()
        self._dependents = set()
        self.setFormula(formula)
        self._is_supporting_cell = supporting_cell

    def __repr__(self):
        return 'SpreadSheetCell(%s, %s)' % (repr(self._name), 
                                            repr(self._formula))

    def getName(self):
        return self._name

    def _removePrecedents(self):
        for precedent in self._precedents:
            precedent._dependents.remove(self)
            if precedent.isUnusedSupportingCell():
                precedent._removePrecedents()
        self._precedents.clear()

    def setFormula(self, formula):
        self._compiledformula = None
        self._formula = formula
        self._removePrecedents()
        self._markAsNotUpdated()
        self._is_supporting_cell = False

    def isUnusedSupportingCell(self):
        return self._is_supporting_cell and not self._dependents

    def isSupportingCell(self):
        return self._is_supporting_cell

    def getFormula(self):
        return self._formula

    def getValue(self):
        #if not self._updated:
            #import traceback
            #print self._formula
            #traceback.print_stack()
            #print 
            #print
        if not self._updated:
            if not self._compiledformula:
                self._compiledformula = compiler.compile(self._formula, 'Cell formula', 'eval')
            self._cashedvalue = self._evaluator.eval(self._compiledformula)
            self._updated = True
        return self._cashedvalue 

    def getPrecedents(self, include_supporting=True):
        if include_supporting:
            return self._precedents
        else:
            return self._getTrueDependendencies(SpreadSheetCell.getPrecedents)

    def getDependents(self, include_supporting=True):
        if include_supporting:
            return self._dependents
        else:
            return self._getTrueDependendencies(SpreadSheetCell.getDependents)

    def _getTrueDependendencies(self, f):
        """ Remove suporting cells from set of precedents/dependents """
        true_deps = set()
        for d in f(self):
            if d.isSupportingCell():
                true_deps.update(f(d))
            else:
                true_deps.add(d)
        return true_deps

    def _markAsNotUpdated(self):
        if self._updated:
            for dependent in self._dependents:
                dependent._markAsNotUpdated()
            self._updated = False

    def addDependent(self, dependent):
        self._dependents.add(dependent)
        dependent._precedents.add(self)

    def remove(self):
        self._removePrecedents()

class SpreadSheet:
    def __init__(self, additional_paths=[], case_sensitive_names=False):
        self._cells = {}
        self._locals = {}
        self._globals = {}
        self._script = ''
        self._additional_paths = additional_paths
        #What cells I am evaluating right now
        self._eving = []
        #regex_compile_flags = re.IGNORECASE if not case_sensitive_names else None
        regex_compile_flags = 0
        self._cell_name_regex = re.compile('^(?P<col>[a-z])(?P<row>\d+)$', 
			regex_compile_flags)
        relative_name_regex = r"""
^
( # First part, column name
 (?P<col>[a-z]) # An absolute column
 |
 (?P<relcol>X
  (
   (?P<coloffset>\d+) # Offset
   (?P<coloffsetsign>[mplr]) # Plus / minus / left / right
  )?
 )
)
( # Second part, row name
 (?P<row>\d+)   # An absolute row
 |
 (?P<relrow>X
  (
   (?P<rowoffset>\d+) # Offset
   (?P<rowoffsetsign>[mpud]) # Plus / minus / up / down
  )?
 )
)
$
"""
        self._rel_cell_name_regex = re.compile(relative_name_regex, re.VERBOSE)
        #self._rel_cell_name_regex = (
                        #re.compile('^(?P<col>[a-zX])(?P<row>\d+|X)$',
			#regex_compile_flags))
        self._cell_range_regex = re.compile(
			'^(?P<start_column>[a-z])(?P<start_row>\d+)' +
			'_to_' +
			'(?P<end_column>[a-z])(?P<end_row>\d+)$', 
			regex_compile_flags)

    def _calculate(self):
        # Can not use iterkeys, because some supporting cells may be added. 
        # They will be calculated, so we do not have to iterate over them again.
        for cell_name in self._cells.keys():
            if self._cells[cell_name].isUnusedSupportingCell():
                del self._cells[cell_name]
            else:
                _ = self[cell_name]

    def isEmptyCell(self, cell_name):
        return cell_name not in self._cells

    def getCell(self, cell_name):
        cell = self._cells.get(cell_name, None)
        if cell:
            # TODO There is something horribly wrong with the relationship spreadsheet <-> cell <-> evaluator
            try:
                _ = self[cell_name] # Value must be calculated, otherwise dependencies not correct
            except:
                pass
        return cell

    def setCellFormula(self, cell_name, formula):
        self[cell_name] = formula

    def deleteCell(self, cell_name):
        if cell_name in self._cells:
            self._cells[cell_name].remove()
            del self._cells[cell_name]

    def setScript(self, script):
        self._script = script
        self._locals.clear()
        self._globals.clear()
        for cell in self._cells.itervalues():
            cell._markAsNotUpdated()

        import sys
        original_path = sys.path[:]
        sys.path.extend(self._additional_paths)
        try:
            # I can't put the defined functions into _locals, 
            # they will not show in other functions then... 
            exec(self._script, self._globals, self._globals)
        except:
            sys.path[:] = original_path[:]
            raise

    def getScript(self):
        return self._script

    def save(self):
        format = \
"""
cells = \
{
%s
}

script = \
%s
"""
        sorted_cell_names = sorted([(self.getCellNamePos(name), name, cell )
                                   for name, cell in self._cells.items()])

        cells = ('    %s: %s' % (repr(cell_name), repr(cell.getFormula())) 
                  for _, cell_name, cell in sorted_cell_names 
                  if not cell.isSupportingCell())
        cells = ', \n'.join(cells)
        return format % (cells, repr(self._script))

    def load(self, file):
        locals = {}
        exec(file, {}, locals)
        cells = locals['cells']
        self._cells.clear()
        for name, formula in cells.iteritems():
            self[name] = formula
        self.setScript(locals['script'])

    def clear(self):
        self._cells.clear()
        self.setScript('')

    def getCellInfo(self, cell):
        name = cell.getName()
        formula = cell.getFormula()
        precedents = len(cell.getPrecedents())
        dependents = len(cell.getDependents())
        if cell.isSupportingCell():
            label = None
        else:
            row, col = self.getCellNamePos(name)
            label_cell_name = self.getCellName(row, col-1)
            if label_cell_name is not None:
                label_cell = self.getCell(label_cell_name)
                label = label_cell.getValue() if label_cell else None
                if not isinstance(label, str):
                    label = None
            else:
                label = None
        return name, label, formula, precedents, dependents

    def asScript(self, print_results=False):
        self._calculate()
        script = ['#!/usr/bin/python', 
                  '## Generated by Python Scriptable Spread Sheet', 
                  '', 
                  '# The init script', 
                  self._script, 
                  '',
                  '# Formulas from the spread sheet',
                  #'def f():', 
                  ]
        cell_precedent_count = dict((c, len(c.getPrecedents())) 
                                    for c in self._cells.itervalues())
        while cell_precedent_count:
            unprecedented = [c for c, n in cell_precedent_count.iteritems() 
                             if n == 0]
            for cell in unprecedented:
                name, label, formula, precedents, dependents = self.getCellInfo(cell)
                if dependents + precedents > 0:
                    if label:
                        script.append('# %s' % label)
                    script.append('%s = %s' % (name, formula))
                for d in cell.getDependents():
                    cell_precedent_count[d] -= 1
                del cell_precedent_count[cell]
        if print_results:
            script.extend(
            ['',
             '# Show some results',
             r"print '\n'.join('%s: %s' % d for d in locals().iteritems() ",
             "                if not d[0].startswith('__'))"])

        return '\n'.join(script) + '\n'

    def asDot(self):
        self._calculate()
        dot = ['## Generated by Python Scriptable Spread Sheet',
               'digraph spreadsheet {',
               '',
                 ]

        nodes = []
        for n, c in self._cells.iteritems() :
            name, cell_label, formula, precedents, dependents = \
                self.getCellInfo(c)
            if not c.isSupportingCell() and (dependents or precedents):
                value = str(c.getValue())
                label = []
                if cell_label is not None:
                    label.append(cell_label)
                if "'" + value + "'" != formula:
                    label.append(formula)
                label.append('=%s' % (value))
                label = r'\n'.join(label)
                nodes.append('%s [shape=box, label="%s"]' % (n, label))

        edges = []
        for n, c in self._cells.iteritems():
            for dep in c.getDependents(False):
                edges.append('%s->%s [label="%s"]' % (n, dep.getName(), n))
        dot.extend(nodes)
        dot.append('')
        dot.extend(edges)
        dot.append('}')
        return '\n'.join(dot) + '\n'

    def __delitem__(self, key):
        if key in self._cells:
            del self._cells[key]

    def __setitem__(self, key, formula):
        if re.match(self._cell_name_regex, key):
            formula = formula.strip()
            if key in self._cells:
                self._cells[key].setFormula(formula)
            else:
                self._cells[key] = SpreadSheetCell(key, formula, self)
        else:
            # For instance list comprehension use 'special' variables
            self._locals[key] = formula

    def eval(self, compiledformula):
        return eval(compiledformula, self._globals, self)

    def _columnPos2Name(self, pos):
        return chr(ord('a') + pos)

    def _rowPos2Name(self, pos):
        return str(pos + 1)

    def _columnName2Pos(self, pos):
        return ord(pos) - ord('a')

    def _rowName2Pos(self, pos):
        return int(pos) - 1

    def getCellName(self, row, col):
        if (row < 0 or col < 0):
            return None
        return chr(ord('a') + col) + str(row + 1)

    def getCellNamePos(self, name):
        name_match = re.match(self._cell_name_regex, name)
        if not name_match:
            return None
        row = name_match.group('row')
        col = name_match.group('col')
        return (int(row) - 1, ord(col) - ord('a'))

    def _makeRelativeCellReferenceAbsolute(self, relative_ref, base_name):
        relative_match = re.match(self._rel_cell_name_regex, relative_ref)
        if not relative_match:
            return None
        else:
            md = relative_match.groupdict()
            base_row, base_col = self.getCellNamePos(base_name)
            signs = {
                'p': 1,
                'r': 1,
                'd': 1,
                'm': -1,
                'l': -1,
                'u': -1,
                None: 0
            }
            if md['relcol']:
                relcol = signs[md['coloffsetsign']] * int(md['coloffset'] or 0)
                col = base_col + relcol
            else:
                col = self._columnName2Pos(md['col'])
            if md['relrow']:
                relrow = signs[md['rowoffsetsign']] * int(md['rowoffset'] or 0)
                row = base_row + relrow
            else:
                row = self._rowName2Pos(md['row'])
            cell_name = self.getCellName(row, col)
            if cell_name is None:
                raise SpreadSheetError('Relative cell reference outside bounds')
            return cell_name

    def _getRangeFormula(self, range_name):
        def column_name_to_number(name):
            return ord(name)-ord('a')

        def column_number_to_name(n):
            return chr(n+ord('a'))

        def row_range(row, left, right):
            return '[%s]' % ', '.join(column_number_to_name(c)+str(row)
                                      for c in xrange(left, right+1))

        def column_range(top, bottom, column):
            return '[%s]' % ', '.join(column_number_to_name(column)+str(r)
                                      for r in xrange(top, bottom+1))

        range_match = re.match(self._cell_range_regex, range_name)
        if not range_match:
            return None
        elif range_match:
            left = column_name_to_number(range_match.group('start_column'))
            right = column_name_to_number(range_match.group('end_column'))
            top = int(range_match.group('start_row'))
            bottom = int(range_match.group('end_row'))
            if top == bottom:
                return row_range(top, left, right)
            elif left == right:
                return column_range(top, bottom, left)
            else:
                return '[%s]' % ', '.join(row_range(r, left, right)
                                          for r in xrange(top, bottom+1))

    def __getitem__(self, key):
        is_outmost_call = not self._eving

        try:
            if key in self._cells:
                cell = self._cells[key]
                if cell in self._eving:
                    #Evaluating a cell in a loop
                    del self._eving[:]
                    raise SpreadSheetError('Circular Dependency')

                if self._eving:
                    cell.addDependent(self._eving[-1])

                self._eving.append(cell)
                rv = cell.getValue()
                self._eving.pop()
            else:
                range_formula = self._getRangeFormula(key)
                if re.match(self._cell_name_regex, key):
                    # This is a valid cell name, but that cell is empty
                    self._cells[key] = SpreadSheetCell(key, repr(None), self, 
                                                       supporting_cell=True)
                    rv = self[key]
                elif range_formula:
                    # This is a valid range name, so create a new supporting
                    # cell
                    self._cells[key] = SpreadSheetCell(key, range_formula, self,
                                                       supporting_cell=True)
                    rv = self[key]
                else:
                    base_cell = self._eving[-1] if self._eving else None
                    cell_name = self._makeRelativeCellReferenceAbsolute(key, base_cell.getName()) if base_cell else None
                    if cell_name:
                        rv = self.__getitem__(cell_name)
                    else:
                        rv = self._locals[key]
        finally:
            if is_outmost_call:
                #Regardless of wheter an exception was raised, clear _eving
                del self._eving[:]

        return rv


def main():
    import unittest
    class TestSpreadSheetBasics(unittest.TestCase):

        def testCellNames(self):
            ss = SpreadSheet()
            ss.setCellFormula('a1', '1')
            ss['z1'] = '2'
            ss['a1000'] = '3'
            self.assertEqual(ss['a1'], 1)
            self.assertEqual(ss['z1'], 2)
            self.assertEqual(ss['a1000'], 3)

        def testFormulasAreEvaluated(self):
            ss = SpreadSheet()
            formula = '3*4'
            ss['a1'] = formula
            self.assertEqual(ss['a1'], 12)
            self.assertEqual(ss.getCell('a1').getValue(), 12)
            self.assertEqual(ss.getCell('a1').getFormula(), formula)

        def testGetCell(self):
            ss = SpreadSheet()
            ss.setCellFormula('a1', '2')
            self.assertEqual(ss.getCell('a1').getFormula(), '2')
            self.assertEqual(ss.getCell('a2'), None)

        def testReferencesCanBeUsed(self):
            ss = SpreadSheet()
            ss['a1'] = '2'
            ss['a2'] = '3*a1'
            self.assertEqual(ss['a2'], 6)

        def testFormulasAreRecalculated(self):
            ss = SpreadSheet()
            ss['a1'] = '2'
            ss['a2'] = '3*a1'
            self.assertEqual(ss['a2'], 6)
            ss['a1'] = '3'
            self.assertEqual(ss['a2'], 9)

        def testSimpleDependenciesAreCalculated(self):
            ss = SpreadSheet()
            ss['a1'] = '2'
            ss['a2'] = '3*a1'
            ss._calculate()
            self.assertEqual(ss.getCell('a1').getDependents(), set([ss.getCell('a2')]))
            self.assertEqual(ss.getCell('a2').getPrecedents(), set([ss.getCell('a1')]))

        def testSimpleDependenciesAreUpdated(self):
            ss = SpreadSheet()
            ss['a1'] = '2'
            ss['a2'] = '3*a1'
            ss._calculate()
            self.assertEqual(ss.getCell('a1').getDependents(), set([ss.getCell('a2')]))
            self.assertEqual(ss.getCell('a2').getPrecedents(), set([ss.getCell('a1')]))
            ss['a2'] = '3'
            ss._calculate()
            self.assertEqual(ss.getCell('a1').getDependents(), set())
            self.assertEqual(ss.getCell('a2').getPrecedents(), set())

        def testDependenciesAreCalculated(self):
            ss = SpreadSheet()
            ss['a1'] = '1'
            ss['a2'] = '2'
            ss['a3'] = '3'
            ss['a4'] = '4'
            ss['b1'] = 'a1*a2'
            ss['b2'] = 'a3*a4'
            ss['c1'] = 'b1*b2'
            ss._calculate()
            self.assertEqual(ss.getCell('c1').getPrecedents(), set(ss.getCell(n) for n in ['b1', 'b2']))
            self.assertEqual(ss.getCell('c1').getDependents(), set())
            self.assertEqual(ss['c1'], 24)

        def testCircularDependenciesAreErrors(self):
            ss = SpreadSheet()
            ss['a1'] = 'a2'
            ss['a2'] = 'a1'
            f = lambda: ss['a1']
            self.assertRaises(SpreadSheetError, f)
            try:
                f()
            except SpreadSheetError, e:
                self.assertEqual(str(e), 'Error: Circular Dependency')

        def testNameErrorsAreReported(self):
            ss = SpreadSheet()
            ss['a1'] = '1'
            ss['b2'] = 'a1+elefant'
            self.assertRaises(NameError, lambda: ss['b2'])
            try:
                val = ss['b2']
            except NameError, e:
                self.assertEqual(str(e), "name 'elefant' is not defined")

        def testUnassignedCellsAreEmpty(self):
            ss = SpreadSheet()
            self.assertEqual(ss['a1'], None)

        def testUpdatedCellsAreTracked(self):
            ss = SpreadSheet()
            ss['a1'] = '1'
            ss['a2'] = '2'
            ss['a3'] = '3'
            ss['a4'] = '4'
            ss['b1'] = 'a1*a2'
            ss['b2'] = 'a3*a4'
            ss['c1'] = 'b1*b2'
            self.assertEqual(ss['c1'], 24)
            self.assertEqual(ss._cells['b1']._updated, True)
            ss['b1'] = '2*a1*a2'
            self.assertEqual(ss._cells['a1']._updated, True)
            self.assertEqual(ss._cells['b1']._updated, False)
            self.assertEqual(ss._cells['c1']._updated, False)

        def testFormulasCanUseListComprehension(self):
            ss = SpreadSheet()
            ss.setCellFormula('a1', '[x for x in xrange(3)]')
            self.assertEqual(ss.getCell('a1').getValue(), [0, 1, 2])

        def testColumnRangesAreLists(self):
            ss = SpreadSheet()
            ss['a1'] = repr('A')
            ss['b1'] = repr('B')
            ss['c1'] = repr('C')
            self.assertEqual(ss['a1_to_c1'], ['A', 'B', 'C'])

        def testRowRangesAreLists(self):
            ss = SpreadSheet()
            ss['a1'] = '1'
            ss['a2'] = '2'
            ss['a3'] = '3'
            self.assertEqual(ss['a1_to_a3'], [1, 2, 3])

        def testRectangularRangesWork(self):
            ss = SpreadSheet()
            ss['a1'] = repr('a1')
            ss['a2'] = repr('a2')
            ss['a3'] = repr('a3')
            ss['b1'] = repr('b1')
            ss['b2'] = repr('b2')
            ss['b3'] = repr('b3')
            self.assertEqual(ss['a1_to_b3'], [['a1', 'b1'], ['a2', 'b2'], ['a3', 'b3']])

        def testRelativeColumnDependenciesWork(self):
            ss = SpreadSheet()
            ss['a1'] = repr(1)
            ss['a2'] = 'a1'
            ss['a3'] = 'X1'
            self.assertEqual(ss['a3'], 1)

        def testRelativeRowDependenciesWork(self):
            ss = SpreadSheet()
            ss['a1'] = repr(1)
            ss['b1'] = 'a1'
            ss['c1'] = 'aX'
            self.assertEqual(ss['c1'], 1)

        def testRelativeDependenciesWork(self):
            ss = SpreadSheet()
            ss['b1'] = repr(2)
            ss['b5'] = 'XX4u'
            self.assertEqual(ss['b5'], 2)

        def testRelativeDependenciesReferingOutsideSheetWork(self):
            ss = SpreadSheet()
            ss['b1'] = 'bX1u'
            ss['a2'] = 'X1l2'
            f = lambda: ss['b1']
            self.assertRaises(SpreadSheetError, f)
            try:
                f()
            except SpreadSheetError, e:
                self.assertEqual(str(e), 'Error: Relative cell reference outside bounds')
            f = lambda: ss['a2']
            self.assertRaises(SpreadSheetError, f)

        def testScriptsVariablesCanBeUsedInCells(self):
            ss = SpreadSheet()
            ss.setScript('ONE_THOUSAND = 1000')
            ss['a1000'] = 'ONE_THOUSAND'
            self.assertEqual(ss['a1000'], 1000)

        def testFunctionsInScriptCanUseOtherFunctionsInScript(self):
            ss = SpreadSheet()
            ss.setScript("""
def a():
    return 1

def b():
    return 2*a()
""")
            ss['a1'] = 'a()'
            ss['b1'] = 'b()'
            self.assertEqual(ss['a1'], 1)
            self.assertEqual(ss['b1'], 2)

        def testSettingScriptsRecalculatesCells(self):
            ss = SpreadSheet()
            ss.setScript('ONE_THOUSAND = 1000')
            ss['a1000'] = 'ONE_THOUSAND'
            self.assertEqual(ss['a1000'], 1000)
            ss.setScript('ONE_THOUSAND = 900') # Inflation
            self.assertEqual(ss['a1000'], 900)

        def testCellsCanUseFunctionsDefinedInScripts(self):
            ss = SpreadSheet()
            ss.setScript(
"""
def squared(n):
    return n*n
""")
            ss['b4'] = 'squared(4)'
            self.assertEqual(ss['b4'], 16)

        def testScriptsCanImportStuff(self):
            ss = SpreadSheet()
            ss.setScript(
"""
from math import log10
""")
            ss['a1'] = 'log10(100)'
            self.assertEqual(ss['a1'], 2.0)

        def testSavingEmptySpreadSheet(self):
            ss = SpreadSheet()
            save_file = \
"""
cells = \
{

}

script = \
''
"""
            self.assertEquals(ss.save(), save_file)

        def testSavingSpreadSheet(self):
            ss = SpreadSheet()
            ss['a1'] = '10'
            ss['a2'] = '100'
            ss['b1'] = 'log10(a1)'
            ss['b2'] = 'log10(a2)'
            ss.setScript('from math import log10')
            save_file = \
"""
cells = \
{
    'a1': '10', 
    'b1': 'log10(a1)', 
    'a2': '100', 
    'b2': 'log10(a2)'
}

script = \
'from math import log10'
"""
            self.assertEquals(ss.save(), save_file)

        def testLoadingSpreadSheet(self):
            ss = SpreadSheet()
            save_file = \
"""
cells = \
{
    'a1': '10', 
    'b1': 'log10(a1)', 
    'a2': '100', 
    'b2': 'log10(a2)'
}

script = \
'from math import log10'
"""
            ss.load(save_file)
            self.assertEquals(ss['b2'], 2.0)
            self.assertEquals(ss.save(), save_file)

        def testAsScript(self):
            ss = SpreadSheet()
            ss['a1'] = '10'
            ss['b1'] = 'log10(a1)'
            ss['a2'] = 'log10(b2)'
            ss['b2'] = '100'
            ss.setScript('from math import log10')
            save_script = \
"""#!/usr/bin/python
## Generated by Python Scriptable Spread Sheet

# The init script
from math import log10

# Formulas from the spread sheet
a1 = 10
b2 = 100
a2 = log10(b2)
b1 = log10(a1)

# Show some results
print '\\n'.join('%s: %s' % d for d in locals().iteritems() 
                if not d[0].startswith('__'))
"""
            self.assertEquals(ss.asScript(print_results=True), save_script)
            d = {}
            exec ss.asScript() in d
            self.assertEquals(d['a2'], 2.0)

        def testAsDot(self):
            ss = SpreadSheet()
            ss['a2'] = repr('sum')
            ss['a1'] = '5'
            ss['b1'] = '6'
            ss['b2'] = 'a1+b1+10'
            save_dot = \
"""## Generated by Python Scriptable Spread Sheet
digraph spreadsheet {

a1 [shape=box, label="5\\n=5"]
b1 [shape=box, label="6\\n=6"]
b2 [shape=box, label="sum\\na1+b1+10\\n=21"]

a1->b2 [label="a1"]
b1->b2 [label="b1"]
}
"""
            self.assertEquals(ss.asDot(), save_dot)

        def testScriptCombinesWithRanges(self):
            ss = SpreadSheet()
            ss['a1'] = repr('a')
            ss['a2'] = '2'
            ss['a3'] = repr(None)
            ss['b1'] = repr('A label for c1')
            ss['c1'] = 'a1_to_a4'
            d = {}
            exec ss.asScript() in d
            self.assertEquals(d['a1'], 'a')
            self.assertEquals(d['a3'], None)
            self.assertEquals(d['a1_to_a4'], ['a', 2, None, None])
            self.assertEquals(d['c1'], ['a', 2, None, None])

        def testUnusedSupportCellsAreRemoved(self):
            ss = SpreadSheet()
            ss['a1'] = 'a2'
            ss['a3'] = 'b1_to_c3'
            d = {}
            exec ss.asScript() in d
            self.assertEquals(d['a1'], None)
            self.assertEquals(d['a2'], None)
            self.assert_('a4' not in d)
            self.assert_('b1_to_c3' in d)

            ss['a1'] = '1'
            ss['a3'] = 'a1'
            d = {}
            exec ss.asScript() in d
            self.assertEquals(d['a1'], 1)
            self.assert_('a2' not in d)
            self.assert_('b1_to_c3' not in d)

        def testGetCellName(self):
            ss = SpreadSheet()
            self.assertEquals(ss.getCellName(0, 0), 'a1')
            self.assertEquals(ss.getCellName(99, 0), 'a100')
            self.assertEquals(ss.getCellName(99, 25), 'z100')
            self.assertEquals(ss.getCellName(-1, 5), None)

        def testGetCellNamePos(self):
            ss = SpreadSheet()
            self.assertEquals(ss.getCellNamePos('a1'), (0, 0))
            self.assertEquals(ss.getCellNamePos('z100'), (99, 25))

        def testMakeRelativeReferenceAbsolute(self):
            ss = SpreadSheet()
            make_relative = ss._makeRelativeCellReferenceAbsolute
            self.assertEquals(make_relative('X1lX', 'c20'), 'b20')
            self.assertEquals(make_relative('X1r30', 'c20'), 'd30')
            self.assertEquals(make_relative('X1mX', 'c20'), 'b20')
            self.assertEquals(make_relative('X1p3', 'c20'), 'd3')
            self.assertEquals(make_relative('dX1u', 'c20'), 'd19')
            self.assertEquals(make_relative('dX1d', 'c20'), 'd21')
            self.assertEquals(make_relative('dX1m', 'c20'), 'd19')
            self.assertEquals(make_relative('dX1p', 'c20'), 'd21')
            f = lambda: make_relative('X1l5', 'a1')
            self.assertRaises(SpreadSheetError, f)

    suite = unittest.TestLoader().loadTestsFromTestCase(TestSpreadSheetBasics)
    #suite = unittest.TestSuite([TestSpreadSheetBasics('testFunctionsInScriptCanUseOtherFunctionsInScript')])
    unittest.TextTestRunner(verbosity=2).run(suite)

if __name__ == '__main__':
    main()
