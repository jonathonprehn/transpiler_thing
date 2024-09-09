


# functions for taking the formula ast's and generating python functions 
# from them, and also other memory management stuff 

# maintains info for supporting code generation
class ProgramInfo:

    def __init__(self):
        self.formula_name_to_function_name = dict()
        self.program_constants = dict()

    def set_func_name_for(self, sheet, name, fnc):
        self.formula_name_to_function_name[(sheet, name)] = fnc
 
    def get_func_name_for(self, sheet, name):
        return self.formula_name_to_function_name[(sheet, name)]

    def get_const(self, sheet, name):
        return self.program_constants[(sheet, name)]
    
    def define_const(self, sheet, name, value):
        self.program_constants[(sheet, name)] = value 



def formula_to_python_function(formula_obj, program_info):
    
    formula_id = formula_obj["formula_id"]
    sheet = formula_obj["sheet"]
    name = formula_obj["name"]
    formula_txt = formula_obj["formula"]
    nodes = formula_obj["parsed"]
    python_function_name = program_info.get_func_name_for(sheet, name)

    # analyze the variables that are involved in the nodes and generate a parameter list 
    param_list = []

    func_lines = []
    func_lines.append("# formula " + str(formula_id))
    func_lines.append("# sheet = " + str(sheet))
    func_lines.append("# name = " + str(name))
    func_lines.append("# Excel formula:")
    func_lines.append("# " + str(formula_txt))
    func_lines.append("def " + str(python_function_name) + "(" + ",".join(param_list) + ")")
    func_lines.append("    pass")
    func_lines.append("")
    func_lines.append("")

    return "\n".join(func_lines)    
    

