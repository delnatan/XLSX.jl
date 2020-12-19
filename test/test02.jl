# testing multicolumn read

import XLSX
using DataFrames

multicolumns = XLSX.multicolumn_range("B,D:F")

out = XLSX.readtable(
    "/Users/delnatan/Repository/XLSX.jl/data/multicolumn.xlsx", 
    "Sheet1", 
    multicolumns;
    header=true,
    infer_eltypes=true
    )