
import os 
import sys 

from .gen import ExcelScanResults 

from typing import Iterable
from lexit import Lexer, Token


# the lexical analyzer will be separate from this
# but it will be output in a form where
# the python part can be constructed 



# each formula needs a source and a target
# memory location. The program is an 
# ordered set of formulas with growing and 
# shrinking memory according to how many
# free cells there are. 

class IRNode:

    def __init__(self):
        self.parent = None 

    def get_parent(self):
        return self.parent 
    


class IRConstant(IRNode):

    def __init__(self, value):
        super()
        self.value = value

    def value(self):
        return self.value 
    


class IRVariable(IRNode):

    def __init__(self, symbol):
        super()
        self.varname = symbol 
        self.sheet_scope = "$$$GLOBAL$$$"
    
    def sheet_scope(self):
        return self.sheet_scope

    def varname(self):
        return self.varname
    
    def __eq__(self, other):
        return other and self.varname == other.varname and self.sheet_scope == other.sheet_scope 

    def __hash__(self):
        return hash((self.varname, self.sheet_scope))

    def __ne__(self, other):
        return not self.__eq__(other)

    def __str__(self):
        return str(self.sheet_scope) + ":" + str(self.varname)




class UnaryOperationNode(IRNode):

    def __init__(self, op, node):
        super()
        self.op = op 
        self.node = node 

    def node(self):
        return self.node

    def operator(self):
        return self.op 
    


class BinaryOperationNode(IRNode):

    def __init__(self, l, op, r):
        super()
        self.left = l 
        self.op = op 
        self.right = r

    def left(self):
        return self.left
    
    def right(self):
        return self.right 
    
    def operator(self):
        return self.op 
   


class FunctionCallNode(IRNode):

    def __init__(self, func_name, params):
        super()
        self.func_name = func_name 
        self.params = params 

    def func_name(self):
        return self.func_name 
    
    def params(self):
        return self.params


# need to post process the tokens for handling strings, 
# prob not setup the best 
class ExcelFormulaLexer(Lexer):
    NUMBER = r'-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?'
    SINGLE_QUOTE = '\''
    DBL_QUOTE = '\"'
    CELLNAME = r'(\$?)[A-Za-z]([A-Za-z])*(\$?)([0-9])+'
    NAME = r'[A-Za-z_]([A-Za-z_])*'
    WHITESPACE = r'\s+'
    L_BRACE = r'{'
    R_BRACE = r'}'
    L_BRACKET = r'\['
    R_BRACKET = r'\]'
    TRUE = r'TRUE'
    FALSE = r'FALSE'
    LPAREN = '\('
    RPAREN = '\)'
    EXCLAIMATION_POINT = '\!'
    COMMA = r','
    COLON = r':'
    ADD = '\+'
    SUB = '-'
    MUL = '\*'
    DIV = '/'
    EQUALS_SIGN = '='
    LESS_THAN_EQUAL = '<='
    LESS_THAN = '<'
    GREATER_THAN_EQUAL = '>='
    GREATER_THAN = '>'
    CARROT = '\^'
    AMPERSAND = '\&'
    REF= "!#REF!"
    DOLLAR_SIGN = '\$'



# convert the excel scanned formulas
# to an intermediate representation

def serialize_IR_node():
    pass



def excel_formula_to_IR(excel_formula):
    # print(excel_formula)
    
    tokens_iter = list(ExcelFormulaLexer.lex(excel_formula))
    # print(*tokens_iter, sep="\n")

    # post process the tokens

    if tokens_iter[0].type == "EQUALS_SIGN":
        print('formula')

        for token in tokens_iter[1:]:
            if token is not None:
                t = token.type 
                v = token.value 
                print(str(v) + " : " + str(t))
        

    else:
        # this will probably never be seen 
        print('constant')
        print(tokens_iter[1:])
    

    




def materialize_ir_to_python(node, target_function):
    pass 



