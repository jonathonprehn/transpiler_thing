
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
    
    def __str__(self):
        return "IRConstant<" + str(self.value) + ">"
    


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
        return "IRVariable<" + str(self.sheet_scope) + "::" + str(self.varname) + ">"




class UnaryOperationNode(IRNode):

    def __init__(self, op, node):
        super()
        self.op = op 
        self.node = node 

    def node(self):
        return self.node

    def operator(self):
        return self.op 
    
    def __str__(self):
        return "UnaryOperationNode<" + str(self.op) + ">"
    


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
    
    def __str__(self):
        return "BinaryOperationNode<" + str(self.op) + ">"
   


class FunctionCallNode(IRNode):

    def __init__(self, func_name, params):
        super()
        self.func_name = func_name 
        self.params = params 

    def func_name(self):
        return self.func_name 
    
    def params(self):
        return self.params
    
    def __str__(self):
        return "FunctionCallNode<" + str(self.func_name) + "(" + ", ".join(self.params) + ")>"


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


class Stack:
    def __init__(self):
        self.stk = []
        self.ptr = -1
    
    def push(self, obj):
        self.stk.append(obj)
        self.ptr = self.ptr + 1

    def pop(self):
        if self.ptr >= 0:
            ret = self.stk[self.ptr]
            del self.stk[self.ptr]
            self.ptr = self.ptr - 1
            return ret
        else:
            return None

    def peek(self):
        if self.ptr >= 0:
            return self.stk[self.ptr]
        else:
            return None 
        
    def size(self):
        return len(self.stk)


def excel_formula_to_IR(excel_formula):
    # print(excel_formula)
    
    tokens_iter = list(ExcelFormulaLexer.lex(excel_formula))
    # print(*tokens_iter, sep="\n")

    # post process the tokens

    if tokens_iter[0].type == "EQUALS_SIGN":
        print('formula')

        node_stack = Stack()
        token_string_processing_stk = Stack()

        # counters+trackers, this is probably a bad way to do this kind of thing 
        # cur_string = ""
        # within_string_dbl_quote_counter = 0

        ntokens = len(tokens_iter)
        itr = 1
        while itr < ntokens:
            token = tokens_iter[itr]
            if token is not None:
                t = token.type 
                v = token.value 
                
                # print(str(v) + " : " + str(t))

                is_within_string = False 

                if token_string_processing_stk.size() > 0:  # currently processing a string 
                    is_within_string = True

                if is_within_string:

                    def unravel_string():
                        unraveled = ""
                        unpopped = []
                        ending_dbl_quote = token_string_processing_stk.pop()
                        while token_string_processing_stk.size() > 1:
                            next_token = token_string_processing_stk.pop()
                            unpopped.append(next_token)
                        starting_dbl_quote = token_string_processing_stk.pop()

                        assert(ending_dbl_quote.type == "DBL_QUOTE" and starting_dbl_quote.type == "DBL_QUOTE")
                        assert(token_string_processing_stk.size() == 0)
                        # reverse order from the stack to rebuild the string correctly
                        unpopped.reverse()
                        for tok in unpopped:
                            unraveled = unraveled + tok.value
                        
                        return unraveled

                    # if a double quote and the next token is not a double quote, then end the string, 
                    # unravel it, and construct a constant node out of it, 
                    # which you can add to the stack 

                    if t == 'DBL_QUOTE':

                        if itr >= ntokens-1:
                            # end of the formula? - end string (assumes all formulas have valid syntax)
                            token_string_processing_stk.push(token)
                            the_string = unravel_string()
                            str_node = IRConstant(the_string)
                            node_stack.push(str_node)

                        elif tokens_iter[itr+1].type == 'DBL_QUOTE':
                            # is the next one a double quote? - collapse to a single one, do not end the string 

                            itr = itr + 1 # ignore the next double quote
                            token_string_processing_stk.push(token) # don't end the string 
                        else:
                            # the next token after this double quote is not also a double quote, end the string
                            token_string_processing_stk.push(token) 
                            the_string = unravel_string()
                            str_node = IRConstant(the_string)
                            node_stack.push(str_node)

                    else:
                        # still in the string, so just add it to process like normal  
                        token_string_processing_stk.push(token)

                else:
                    if t == 'NUMBER':
                        pass
                    elif t == 'SINGLE_QUOTE':  # just part of a string 
                        pass
                    elif t == 'DBL_QUOTE':
                        
                        if not is_within_string: # redundant, but for readability
                            token_string_processing_stk.push(token)
                        else:
                            raise Exception("Error processing string")

                    elif t == 'CELLNAME':
                        pass
                    elif t == 'NAME':
                        pass
                    elif t == 'WHITESPACE':
                        pass
                    elif t == 'L_BRACE':
                        pass
                    elif t == 'R_BRACE':
                        pass
                    elif t == 'L_BRACKET':
                        pass
                    elif t == 'R_BRACKET':
                        pass
                    elif t == 'TRUE':
                        pass
                    elif t == 'FALSE':
                        pass
                    elif t == 'LPAREN':
                        pass
                    elif t == 'RPAREN':
                        pass
                    elif t == 'EXCLAIMATION_POINT':
                        pass
                    elif t == 'COMMA':
                        pass
                    elif t == 'COLON':
                        pass
                    elif t == 'ADD':
                        pass
                    elif t == 'SUB':
                        pass
                    elif t == 'MUL':
                        pass
                    elif t == 'DIV':
                        pass
                    elif t == 'EQUALS_SIGN':
                        pass
                    elif t == 'LESS_THAN_EQUAL':
                        pass
                    elif t == 'LESS_THAN':
                        pass
                    elif t == 'GREATER_THAN_EQUAL':
                        pass
                    elif t == 'GREATER_THAN':
                        pass
                    elif t == 'CARROT':
                        pass
                    elif t == 'AMPERSAND':
                        pass
                    elif t == 'REF':
                        pass
                    elif t == 'DOLLAR_SIGN':
                        pass
            
            itr = itr + 1

        print("NODE STACK AT THE END")
        for s in node_stack.stk:
            print(s)



    else:
        # this will probably never be seen 
        print('constant')
        print(tokens_iter[1:])
    

    




def materialize_ir_to_python(node, target_function):
    pass 



