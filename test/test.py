
import sys 
import csv 
import traceback 

sys.path.append("../../")

# for p in sys.path:
#     print(p)
    

import transpiler_thing.gen
import transpiler_thing.parse 

input_excel = "simple_formula_testing.xlsx"

result = transpiler_thing.gen.scan_excel(input_excel)
transpiler_thing.gen.dump_scanned_formulas(result, "workspace")
transpiler_thing.gen.dump_scanned_constants(result, "workspace")


# what to determine 
# 1. build graph of references 
# 2. find the leaves 


formulas_csv = transpiler_thing.gen.dump_scanned_formulas_path("workspace")
with open(formulas_csv, "r") as f:
    rdr = csv.DictReader(f)
    for row in rdr:
        formula_id = int(row["formula_id"])
        sheet = row["sheet"]
        name = row["cell_or_name"]
        formula = row["formula"]
        
        try:
            # if formula_id == 106:
            nodes = transpiler_thing.parse.excel_formula_to_IR(formula, in_sheet=sheet)
            # print(nodes)

        except Exception as e:
            print("formula_id = " + str(formula_id))
            print("Exception: " + str(e))
            print(traceback.format_exc())
            quit()



