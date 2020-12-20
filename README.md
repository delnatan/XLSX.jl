
# XLSX.jl

[![License][license-img]](LICENSE)
[![travis][travis-img]][travis-url]
[![appveyor][appveyor-img]][appveyor-url]
[![codecov][codecov-img]][codecov-url]
[![dev][docs-dev-img]][docs-dev-url]
[![stable][docs-stable-img]][docs-stable-url]

[license-img]: http://img.shields.io/badge/license-MIT-brightgreen.svg?style=flat-square
[travis-img]: https://img.shields.io/travis/felipenoris/XLSX.jl/master.svg?logo=travis&label=Linux+/+macOS&style=flat-square
[travis-url]: https://travis-ci.org/felipenoris/XLSX.jl
[appveyor-img]: https://img.shields.io/appveyor/ci/felipenoris/xlsx-jl/master.svg?logo=appveyor&label=Windows&style=flat-square
[appveyor-url]: https://ci.appveyor.com/project/felipenoris/xlsx-jl/branch/master
[codecov-img]: https://img.shields.io/codecov/c/github/felipenoris/XLSX.jl/master.svg?label=codecov&style=flat-square
[codecov-url]: http://codecov.io/github/felipenoris/XLSX.jl?branch=master
[docs-dev-img]: https://img.shields.io/badge/docs-dev-blue.svg?style=flat-square
[docs-dev-url]: https://felipenoris.github.io/XLSX.jl/dev
[docs-stable-img]: https://img.shields.io/badge/docs-stable-blue.svg?style=flat-square
[docs-stable-url]: https://felipenoris.github.io/XLSX.jl/stable

Excel file reader/writer coded in pure Julia.

## Notes for this branch (by delnatan)
I'm learning Julia and wanted to add support for reading data tables from an Excel file with non-contiguous column range (e.g. "A:E,AA,AB,AC:AE"). In addition, I also wanted the reader to have similar behavior with Pandas when reading a table with duplicate header names. This branch is for my experimentation before submitting a pull request. Two additional functions were implemented: multicolumn_range and readtable. 

```
import XLSX

column_ranges = XLSX.multicolumn_range("A:B,M,S,Y,AE,AK")

out = XLSX.readtable(
    "C:\\Users\\delna\\Desktop\\tests\\2020-09-09_rerun-Result.xlsx",
    "Region Detail",
    column_ranges;
    first_row=20,
    header=true,
    infer_eltypes=true
    )
```

## Installation

```julia
julia> Pkg.add("XLSX")
```

## Documentation

Package documentation is hosted at https://felipenoris.github.io/XLSX.jl/stable.

## References

* [ECMA Open XML White Paper](https://www.ecma-international.org/news/TC45_current_work/OpenXML%20White%20Paper.pdf)

* [ECMA-376](https://www.ecma-international.org/publications/standards/Ecma-376.htm)

* [Excel file limits](https://support.office.com/en-gb/article/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3)

## Alternative Packages

* [ExcelFiles.jl](https://github.com/davidanthoff/ExcelFiles.jl)

* [ExcelReaders.jl](https://github.com/davidanthoff/ExcelReaders.jl)

* [XLSXReader.jl](https://github.com/mpastell/XLSXReader.jl)

* [Taro.jl](https://github.com/aviks/Taro.jl)
