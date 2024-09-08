
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
        self.children = [] # this isn't actually used

    def get_parent(self):
        return self.parent 
    
    def add_child(self, c):
        self.children.append(c)

    def children(self):
        return self.children 
    
    def nodetype(self):
        return "node" 
    
    def formatted_str(self):
        return "\{\}"



class IRConstant(IRNode):

    def __init__(self, value):
        super()
        self.value = value
        

    def get_value(self):
        return self.value 
    
    def __str__(self):
        return "IRConstant<" + str(self.value) + ">"
    
    def nodetype(self):
        return "constant"

    
    
# might need to resolve inclusion of $ and have non $ still refer to the same cell 
class IRVariable(IRNode):

    def __init__(self, symbol):
        super()
        self.varname = symbol 
        self.sheet_scope = "$$$GLOBAL$$$"
    
    def get_sheet_scope(self):
        return self.sheet_scope
    
    def set_sheet_scope(self, sht):
        self.sheet_scope = sht

    def get_varname(self):
        return self.varname
    

    def __str__(self):
        return "IRVariable<" + str(self.sheet_scope) + "::" + str(self.varname) + ">"

    def nodetype(self):
        return "variable"
    

# might need to resolve inclusion of $ and have non $ still refer to the same cell 
class IRVariableRange(IRNode):
    
    def __init__(self, symbol1, symbol2):
        super()
        self.varname1 = symbol1
        self.varname2 = symbol2
        self.sheet_scope = "$$$GLOBAL$$$"
    
    def get_sheet_scope(self):
        return self.sheet_scope
    
    def set_sheet_scope(self, sht):
        self.sheet_scope = sht

    def get_varname1(self):
        return self.varname1

    def get_varname2(self):
        return self.varname2

    def __str__(self):
        return "IRVariableRange<" + str(self.sheet_scope) + "::" + str(self.varname1) + ":" + str(self.varname2) + ">"

    def nodetype(self):
        return "variablerange"  # still counts as a variable 



class SymbolNode(IRNode):

    def __init__(self, symbol):
        super()
        self.symbol = symbol

    def get_symbol(self):
        return self.symbol
    
    def __str__(self):
        return "SymbolNode<" + str(self.symbol) + ">"

    def nodetype(self):
        return "symbol"


class UnaryOperationNode(IRNode):

    def __init__(self, op, node):
        super()
        self.op = op 
        self.node = node 

    def get_node(self):
        return self.node

    def set_node(self, n):
        self.node = n 

    def get_operator(self):
        return self.op 
    
    def __str__(self):
        return "UnaryOperationNode<" + str(self.op) + ">"
    
    def nodetype(self):
        return "unary"
    


class BinaryOperationNode(IRNode):

    def __init__(self, l, op, r):
        super()
        self.left = l 
        self.op = op 
        self.right = r

    def get_left(self):
        return self.left
    
    def get_right(self):
        return self.right 
    
    def get_operator(self):
        return self.op 
    
    def __str__(self):
        return "BinaryOperationNode<" + str(self.op) + ">"
   
    def nodetype(self):
        return "binary"



class FunctionCallNode(IRNode):

    def __init__(self, func_name, params):
        super()
        self.func_name = func_name 
        self.params = params 

    def get_func_name(self):
        return self.func_name 
    
    def get_params(self):
        return self.params
    
    def __str__(self):
        return "FunctionCallNode<" + str(self.func_name) + "(" + ", ".join(self.params) + ")>"

    def nodetype(self):
        return "function"



def print_ast_nodes(formula_ast):
    print_ast_nodes_r(formula_ast, 0)


def print_ast_nodes_r(node, lvl):
    
    tab = ""
    for i in range(0, lvl):
        tab = tab + "  "
    
    if node.nodetype() == "constant" \
        or node.nodetype() == "symbol" \
        or node.nodetype() == "variable" \
        or node.nodetype() == "variablerange":
        print(tab + str(node))
    elif node.nodetype() == "binary":
        print_ast_nodes_r(node.get_left(), lvl+1)
        print(tab + str(node.get_operator()))
        print_ast_nodes_r(node.get_right(), lvl+1)
    elif node.nodetype() == "unary":
        print(tab + str(node.get_operator()))
        print_ast_nodes_r(node.get_node(), lvl+1)
    elif node.nodetype() == "function":
        print(tab + str(node.get_func_name()))
        for p in node.get_params():
            print_ast_nodes_r(p, lvl+1)
    else:
        raise Exception("dont have the code to print the type " + str(node.nodetype()))
    
    



# this is not the most correct way to do this, but I referred to these repos while making this:
# - https://github.com/spreadsheetlab/XLParser
# - 
    

# need to post process the tokens for handling strings, 
# prob not setup the best 
class ExcelFormulaLexer(Lexer):
    NUMBER = r'-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?'
    SINGLE_QUOTE = '\''
    DBL_QUOTE = '\"'
    CELLNAME = r'(\$?)[A-Za-z]([A-Za-z])*(\$?)([0-9])+'  # this just ends with numbers instead 
    NAME = r'[A-Za-z_]([A-Za-z_0-9])*'   # this can have letters and numbers intermixed 
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
    NOT_EQUALS_SIGN = '<>'
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


def excel_formula_to_IR(excel_formula, in_sheet="$$$GLOBAL$$$"):
    # print(excel_formula)
    
    tokens_iter = list(ExcelFormulaLexer.lex(excel_formula))
    # print(*tokens_iter, sep="\n")

    # post process the tokens

    if tokens_iter[0].type == "EQUALS_SIGN":
        print('formula   ' + excel_formula)

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
                is_within_single_quote_cell_ref = False

                if token_string_processing_stk.size() > 0 and token_string_processing_stk.stk[0].type == "DBL_QUOTE":  # currently processing a string 
                    is_within_string = True

                if token_string_processing_stk.size() > 0 and token_string_processing_stk.stk[0].type == "SINGLE_QUOTE":  # currently processing a cell reference 
                    is_within_single_quote_cell_ref = True

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
                
                elif is_within_single_quote_cell_ref:

                    # this is going to be similar to the double quote case, only you don't have to handle double single quotes 
                    if t == "SINGLE_QUOTE":
                        # assuming valid syntax, so end this name and unravel the string 
                        token_string_processing_stk.push(token)
                        
                        unraveled = ""
                        unpopped = []
                        ending_single_quote = token_string_processing_stk.pop()
                        while token_string_processing_stk.size() > 1:
                            next_token = token_string_processing_stk.pop()
                            unpopped.append(next_token)
                        starting_single_quote = token_string_processing_stk.pop()

                        assert(ending_single_quote.type == "SINGLE_QUOTE" and starting_single_quote.type == "SINGLE_QUOTE")
                        assert(token_string_processing_stk.size() == 0)
                        # reverse order from the stack to rebuild the string correctly
                        unpopped.reverse()
                        for tok in unpopped:
                            unraveled = unraveled + tok.value

                        str_node = IRConstant(unraveled)
                        node_stack.push(str_node)
                    else:
                        # still in the name, so just add it to process like normal  
                        token_string_processing_stk.push(token)

                else:
                    if t == 'NUMBER':
                        number_node = IRConstant(v)
                        node_stack.push(number_node)
                    elif t == 'SINGLE_QUOTE':  
                        # single quotes outside of strings in Excel formulas mean:
                        # - if it's the first token before =, then it's not a formula, but this case is already handled
                        # - used to treat a cell reference as a single entity (i.e.  'another, sheet'!A5 to mean cell A5 in sheet "another, sheet")
                        
                        # the right strategy is probably to do something like you did for double strings, and collapse them into a single node
                        if not is_within_single_quote_cell_ref: # redundant, but for readability
                            token_string_processing_stk.push(token)
                        else:
                            raise Exception("Error processing string")

                    elif t == 'DBL_QUOTE':
                        
                        if not is_within_string: # redundant, but for readability
                            token_string_processing_stk.push(token)
                        else:
                            raise Exception("Error processing string")

                    elif t == 'CELLNAME':
                        cellname_node = IRVariable(v)
                        node_stack.push(cellname_node)
                    elif t == 'NAME':
                        name_node = IRVariable(v)
                        node_stack.push(name_node)
                    elif t == 'WHITESPACE':
                        pass # ignore whitespace outside of strings
                    elif t == 'L_BRACE':
                        lbrace_node = SymbolNode(v)
                        node_stack.push(lbrace_node)
                    elif t == 'R_BRACE':
                        rbrace_node = SymbolNode(v)
                        node_stack.push(rbrace_node)
                    elif t == 'L_BRACKET':
                        lbracket_node = SymbolNode(v)
                        node_stack.push(lbracket_node)
                    elif t == 'R_BRACKET':
                        rbracket_node = SymbolNode(v)
                        node_stack.push(rbracket_node)
                    elif t == 'TRUE':
                        true_node = IRConstant(True)
                        node_stack.push(true_node)
                    elif t == 'FALSE':
                        false_node = IRConstant(False)
                        node_stack.push(false_node)
                    elif t == 'LPAREN':
                        lparen_node = SymbolNode(v)
                        node_stack.push(lparen_node)
                    elif t == 'RPAREN':
                        rparen_node = SymbolNode(v)
                        node_stack.push(rparen_node)
                    elif t == 'EXCLAIMATION_POINT':
                        exclaimation_pt_node = SymbolNode(v)
                        node_stack.push(exclaimation_pt_node)
                    elif t == 'COMMA':
                        comma_node = SymbolNode(v)
                        node_stack.push(comma_node)
                    elif t == 'COLON':
                        colon_node = SymbolNode(v)
                        node_stack.push(colon_node)
                    elif t == 'ADD':
                        add_node = SymbolNode(v)
                        node_stack.push(add_node)
                    elif t == 'SUB':
                        sub_node = SymbolNode(v)
                        node_stack.push(sub_node)
                    elif t == 'MUL':
                        mul_node = SymbolNode(v)
                        node_stack.push(mul_node)
                    elif t == 'DIV':
                        div_node = SymbolNode(v)
                        node_stack.push(div_node)
                    elif t == 'EQUALS_SIGN':
                        eq_node = SymbolNode(v)
                        node_stack.push(eq_node)
                    elif t == 'NOT_EQUALS_SIGN':
                        eq_node = SymbolNode(v)
                        node_stack.push(eq_node)
                    elif t == 'LESS_THAN_EQUAL':
                        less_than_eq_node = SymbolNode(v)
                        node_stack.push(less_than_eq_node)
                    elif t == 'LESS_THAN':
                        less_than_node = SymbolNode(v)
                        node_stack.push(less_than_node)
                    elif t == 'GREATER_THAN_EQUAL':
                        greater_than_eq_node = SymbolNode(v)
                        node_stack.push(greater_than_eq_node)
                    elif t == 'GREATER_THAN':
                        greater_than_node = SymbolNode(v)
                        node_stack.push(greater_than_node)
                    elif t == 'CARROT':
                        carrot_node = SymbolNode(v)
                        node_stack.push(carrot_node)
                    elif t == 'AMPERSAND':
                        ampersand_node = SymbolNode(v)
                        node_stack.push(ampersand_node)
                    elif t == 'REF':
                        ref_node = IRConstant(v)
                        node_stack.push(ref_node)
                    elif t == 'DOLLAR_SIGN':
                        dollar_node = IRConstant(v)
                        node_stack.push(dollar_node)
                    else:
                        raise Exception("Invalid token type " + str(t))
                        
            
            itr = itr + 1

        # construct syntax tree from the node stack 
        nodes_ordered = list(node_stack.stk)
        # nodes_ordered.reverse()

        # for node in nodes_ordered:
        #     print(node)

        formula_ast = tokens_to_ast_Formula(nodes_ordered, in_sheet=in_sheet)

        if formula_ast is not None:

            print("parsed  ")

            print_ast_nodes(formula_ast)

            print(" ")
            return formula_ast
        else:
            print("Failed to parse formula " + str(excel_formula))
            return None 

    else:
        # this will probably never be seen 
        print('constant')
        print(tokens_iter[1:])
    

    
def tokens_to_ast_Formula(tokens, in_sheet=None):
    ptr = 0

    parsed_node, ptr = tokens_to_ast_Expr(tokens, ptr, in_sheet=in_sheet)
    if parsed_node is not None:
        return parsed_node 

    # nothing matched, what happened?    
    return None 



def tokens_to_ast_Function(tokens, ptr, in_sheet=None):

    print("checking if its a function, ptr = " + str(ptr) + ", token = " + str(tokens[ptr]) )

    func_name = None 
    params = []
    inc_ptr = ptr 
    if inc_ptr < len(tokens) and tokens[inc_ptr].nodetype() == "variable":
        # match function name 
        func_name = tokens[inc_ptr].get_varname()
        inc_ptr = inc_ptr + 1
    else:
        return None, ptr 
    
    print("got function name, is the next symbol a (" )

    # match (
    if inc_ptr < len(tokens) and tokens[inc_ptr].nodetype() == "symbol":
        if tokens[inc_ptr].get_symbol() == '(':
            inc_ptr = inc_ptr + 1
            print("yeah next symbol is (" )
        else:
            print("no next symbol is actually " + str(tokens[inc_ptr].get_symbol()) )
            return None, ptr
    else:
        print("next node is not a symbol at all")
        return None, ptr
    

    # match 0 or more params 

    func_ended = False 

    while not func_ended:

        param_node = None 
        
        if param_node is None:
            param_node, inc_ptr = tokens_to_ast_Expr(tokens, inc_ptr, in_sheet=in_sheet)

        if param_node is not None:
            params.append(param_node)
        else:
            # syntax error 
            print("Cannot match param for a function")
            return None, ptr 

        # match comma or not
        matched_comma = False 
        if inc_ptr < len(tokens) and tokens[inc_ptr].nodetype() == "symbol":
            if tokens[inc_ptr].get_symbol() == ',':
                matched_comma = True  

        # if did not match a comma, need ) 

        # match )
        if not matched_comma:
            if inc_ptr < len(tokens) and tokens[inc_ptr].nodetype() == "symbol":
                if tokens[inc_ptr].get_symbol() == ')':
                    func_ended = True  

        inc_ptr = inc_ptr + 1
        
        if inc_ptr >= len(tokens):
            return None, ptr  # syntax error, didn't end properly

    print("got params")
    for p in params:
        print(p)


    # matched, construct function node and return 
    func_node = FunctionCallNode(func_name, params)

    print("got a function call!, ptr = " + str(inc_ptr) + ", token = " + str(tokens[inc_ptr]) )

    return func_node, inc_ptr


# taking from https://en.wikipedia.org/wiki/Recursive_descent_parser

# thing that is collapsed to add/subtract/multiply/divide/other non-function operators
# + or - a bunch of terms
def tokens_to_ast_Expr(tokens, ptr, in_sheet=None):
    lnode = None 
    rnode = None 
    bin_op = None 
    unary_op = None
    inc_ptr = ptr 

    # print("expr, ptr = " + str(inc_ptr) + ", " + str(tokens[inc_ptr]))
    
    # first - or + for negative or positive, unary operator 
    if inc_ptr < len(tokens) and tokens[inc_ptr].nodetype() == "symbol":
        if tokens[inc_ptr].get_symbol() == '+':
            inc_ptr = inc_ptr + 1
            unary_op = tokens[inc_ptr].symbol()
        elif tokens[inc_ptr].get_symbol() == '-':
            inc_ptr = inc_ptr + 1
            unary_op = tokens[inc_ptr].symbol()
    
    lnode, inc_ptr = tokens_to_ast_Term(tokens, inc_ptr, in_sheet=in_sheet)

    # there is an issue where the factor is not keeping the previous term, skips and doesn't return

    if inc_ptr < len(tokens) and tokens[inc_ptr].nodetype() == "symbol":
        valid_ops = [
            '&', '+', '-', '<', '>',
            '=', '<>', '<=', '>='
        ]
        if inc_ptr < len(tokens) and tokens[inc_ptr].get_symbol() in valid_ops:
            bin_op = tokens[inc_ptr].get_symbol()
            inc_ptr = inc_ptr + 1
            rnode, inc_ptr = tokens_to_ast_Term(tokens, inc_ptr, in_sheet=in_sheet)
            
        else:
            pass # print("invalid binary operator in Expr " + str(tokens[inc_ptr].get_symbol()))
        
    ret_node = None
    if lnode is not None and unary_op is None and rnode is None and bin_op is None:
        # single term only 
        ret_node = lnode
    if lnode is not None and unary_op is not None and rnode is None and bin_op is None:
        # unary only 
        ret_node = UnaryOperationNode(unary_op, lnode)
    elif lnode is not None and unary_op is None and rnode is not None and bin_op is not None:
        # binary op only
        ret_node = BinaryOperationNode(lnode, bin_op, rnode)
    elif lnode is not None and unary_op is not None and rnode is not None and bin_op is not None:
        # binary and unary operations on both ends (i.e. -A1+B2  )
        bin_node = BinaryOperationNode(lnode, bin_op, rnode)
        ret_node = UnaryOperationNode(unary_op, bin_node)

    # print("returning this from expr: " + str(ret_node))
    
    return ret_node, inc_ptr 



# *, / etc
def tokens_to_ast_Term(tokens, ptr, in_sheet=None):
    inc_ptr = ptr 
    
    lfactor = None 
    rfactor = None 
    op = None 

    # print("term, ptr = " + str(inc_ptr) + ", " + str(tokens[inc_ptr]))

    lfactor, inc_ptr = tokens_to_ast_Factor(tokens, inc_ptr, in_sheet=in_sheet)

    # there is an issue where the factor is not keeping the previous term, skips and doesn't return

    # match operator, if it exists 
    if inc_ptr < len(tokens) and tokens[inc_ptr].nodetype() == "symbol":
        # is it a valid symbol?
        valid_ops = [
            '*', '/', '^'
        ]
        if inc_ptr < len(tokens) and tokens[inc_ptr].get_symbol() in valid_ops:
            op = tokens[inc_ptr].get_symbol()
            inc_ptr = inc_ptr + 1
            rfactor, inc_ptr = tokens_to_ast_Factor(tokens, inc_ptr, in_sheet=in_sheet)
        else:
            pass # not an operator for term
        
    else:
        # no operator, might just be a factor by itself 
        pass 

    ret_node = None
    if lfactor is not None and op is None and rfactor is None:
        ret_node = lfactor
    elif lfactor is not None and op is not None and rfactor is not None:
        ret_node = BinaryOperationNode(lfactor, op, rfactor)
    else:
        raise Exception("Did not assemble term properly, vars are \nlfactor = "+str(lfactor)+"\nop = "+str(op)+"\nrfactor = "+str(rfactor))
        
    # print("returning this from term: " + str(ret_node))
    
    return ret_node, inc_ptr

    

# factors are reduced to cell references, expressions, constants or function calls
def tokens_to_ast_Factor(tokens, ptr, in_sheet=None):
    inc_ptr = ptr 

    # print("factor, ptr = " + str(inc_ptr) + ", " + str(tokens[inc_ptr]))

    factor_node = None 
        
    if factor_node is None:
        # print("checking if its a function")
        factor_node, inc_ptr = tokens_to_ast_Function(tokens, inc_ptr)     
    
    if factor_node is None:
        # print("checking if its a cell reference")
        factor_node, inc_ptr = tokens_to_ast_CellRef(tokens, inc_ptr, in_sheet)
    
    if factor_node is None:
        # print("checking if its a constant")
        factor_node, inc_ptr = tokens_to_ast_Constant(tokens, inc_ptr)

    # if factor_node is None:
    #     factor_node, inc_ptr = tokens_to_ast_Expr(tokens, inc_ptr) # infinite recursion if I leave this

    # match (
    if inc_ptr < len(tokens) and tokens[inc_ptr].nodetype() == "symbol":
        # print("checking if its another expression")
        if tokens[inc_ptr].get_symbol() == '(':
            inc_ptr = inc_ptr + 1

            if factor_node is None:
                factor_node, inc_ptr = tokens_to_ast_Expr(tokens, inc_ptr, in_sheet=in_sheet)
            
            
            if inc_ptr < len(tokens) and tokens[inc_ptr].nodetype() == "symbol":
                if tokens[inc_ptr].get_symbol() == ')':
                    inc_ptr = inc_ptr + 1
                else:
                    return None, ptr 
    
    # print("factor_node is " + str(factor_node))

    # not a factor?
    return factor_node, inc_ptr 


# this could also be a range of cells (cell ref : cell ref)
def tokens_to_ast_CellRef(tokens, ptr, in_sheet=None):
    inc_ptr = ptr

    cellname = None 
    sheetname = None 

    is_range = False 

    range_cellname_end = None 

    if inc_ptr < len(tokens) and tokens[inc_ptr].nodetype() == "variable":
        is_variable = True 
        cellname =  tokens[inc_ptr].get_varname()
        inc_ptr = inc_ptr + 1

        # no sheet ref, so implied to be in the same sheet
        sheetname = in_sheet
    else:
        if inc_ptr < len(tokens) and tokens[inc_ptr].nodetype() == "constant":
            is_const = True  
            sheetname = tokens[inc_ptr].get_value()
            inc_ptr = inc_ptr + 1
        else:
            return None, ptr 
        

        if inc_ptr < len(tokens) and tokens[inc_ptr].nodetype() == "symbol":
            if tokens[inc_ptr].get_symbol() == "!":
                pass # ! to indicate sheet name before the symbol
                inc_ptr = inc_ptr + 1
            else:
                return None, ptr 
        else:
            return None, ptr 

        if inc_ptr < len(tokens) and tokens[inc_ptr].nodetype() == "variable":
            is_variable = True  
            cellname =  tokens[inc_ptr].get_varname()
            inc_ptr = inc_ptr + 1
        else:
            return None, ptr 


    # check for : here to see if it's a range! 
    if inc_ptr < len(tokens) and tokens[inc_ptr].nodetype() == "symbol":
        if tokens[inc_ptr].get_symbol() == ":":
            inc_ptr = inc_ptr + 1

            # this is a range, the next symbol must be a variable 

            if inc_ptr < len(tokens) and tokens[inc_ptr].nodetype() == "variable":
                is_variable = True  
                range_cellname_end =  tokens[inc_ptr].get_varname()
                is_range = True
                inc_ptr = inc_ptr + 1
            else:
                return None, ptr 


    # parsed the cell name
    if is_range:
        rangenode = IRVariableRange(cellname, range_cellname_end)
        if sheetname is not None:
            rangenode.set_sheet_scope(sheetname)

        return rangenode, inc_ptr
    else:
        varnode = IRVariable(cellname)
        if sheetname is not None:
            varnode.set_sheet_scope(sheetname)
    
        return varnode, inc_ptr



def tokens_to_ast_Constant(tokens, ptr):
    inc_ptr = ptr
    if inc_ptr < len(tokens) and tokens[inc_ptr].nodetype() == "constant":
        pass 
    else:
        return None, ptr 
    const_node = tokens[inc_ptr]
    inc_ptr = inc_ptr + 1
    return const_node, inc_ptr





def materialize_ir_to_python(node, target_function):
    pass 



