
import sys 
import csv 
import traceback 

sys.path.append("../../")

# for p in sys.path:
#     print(p)
    

import transpiler_thing.gen
import transpiler_thing.parse 
import transpiler_thing.ast_to_python

input_excel = "simple_formula_testing.xlsx"

result = transpiler_thing.gen.scan_excel(input_excel)
transpiler_thing.gen.dump_scanned_formulas(result, "workspace")
transpiler_thing.gen.dump_scanned_constants(result, "workspace")


# what to determine 
# 1. build graph of references 
# 2. find the leaves 


formulas_csv = transpiler_thing.gen.dump_scanned_formulas_path("workspace")
formulas_nodes = transpiler_thing.parse.parse_formulas_csv(formulas_csv)
constants_csv = transpiler_thing.gen.dump_scanned_constants_path("workspace")

programInfo = transpiler_thing.ast_to_python.ProgramInfo()

with open(constants_csv, "r") as cfp:
    drdr = csv.DictReader(cfp)
    for rw in drdr:
        _id = rw["const_id"] 
        sheet = rw["sheet"]
        name = rw["cell_or_name"]
        val = rw["value"]
        # parse the value as a string or a number?
        
        programInfo.define_const(sheet, name, val)


with open("code.py", "w") as codefp:
    
    for formula_obj in formulas_nodes:
        programInfo.set_func_name_for(formula_obj["sheet"], formula_obj["name"], "formula_" + str(formula_obj["formula_id"]))

    for formula_obj in formulas_nodes:
        formula_code = transpiler_thing.ast_to_python.formula_to_python_function(formula_obj, programInfo)
        codefp.write(formula_code)
        codefp.write("\n")
    
