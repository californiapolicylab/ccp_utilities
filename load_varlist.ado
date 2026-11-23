*! load_varlist.ado  version 1.3.0  2026-11-23
*! Load variable lists from an Excel driver file into global macros,
*! using a temporary frame so the dataset in memory is untouched.
*!
*! Syntax:
*!   load_varlist using <filepath> [, sheet(name) varname(colnames)
*!                                   prefix(str) unique verbose ]
*!
*!   The old form  load_varlist, dir(<filepath>) [...]  still works.
*!
*! Options:
*!   sheet(name)         Excel sheet to read; defaults to the first sheet
*!   varname(colnames)   Column(s) to load; default is every column
*!   prefix(str)         Prepend str to each global name, e.g. prefix(vl_)
*!                       turns column "controls" into $vl_controls
*!   unique              Drop repeated entries within each list
*!   verbose             Print each global as it is created
*!
*! Notes:
*!   - One global per column, holding the column's non-blank cells as a
*!     space-separated list. A cell containing spaces contributes
*!     several elements (handy for ranges like "x1-x10 y*").
*!   - Cells containing a backtick, double quote, or $ stop the program
*!     with the offending Excel row, since they break macro expansion.
*!   - Globals are only assigned after every requested column has been
*!     read and checked, so an error never leaves a partial set behind.
*!   - -import excel, firstrow- converts headers to valid Stata names
*!     (e.g., spaces become underscores).
*!
*! Stored results:
*!   r(globals)   names of the globals created
*!
*! History:
*!   1.3.0  using syntax (dir() kept as alias); tempname frame; prefix()
*!          and unique options; special-character check; all-or-nothing
*!          global assignment; stored results; helper program folded in.
*!   1.2.0

program define load_varlist, rclass
    version 16

    syntax [using/] [, DIR(string) SHEET(string) varname(string) ///
        PREfix(string) unique VERBose ]

    // ------------------------------------------------------------------ //
    // Resolve the file path (using, or legacy dir())
    // ------------------------------------------------------------------ //
    if `"`using'"' != "" & `"`dir'"' != "" {
        display as error "Specify either -using- or -dir()-, not both."
        exit 198
    }
    if `"`using'"' == "" & `"`dir'"' == "" {
        display as error "Specify the driver file: load_varlist using <filepath>"
        exit 100
    }
    if `"`using'"' != "" local filepath `"`using'"'
    else local filepath = strtrim(subinstr(`"`dir'"', char(34), "", .))

    capture confirm file `"`filepath'"'
    if _rc {
        display as error `"File not found: [`filepath']"'
        exit 601
    }

    // ------------------------------------------------------------------ //
    // Temporary frame (tempname: unique, dropped automatically on exit)
    // ------------------------------------------------------------------ //
    local caller_frame = c(frame)
    tempname tmp_frame
    frame create `tmp_frame'

    local globals ""

    capture noisily {
        cwf `tmp_frame'

        local sheetopt
        if `"`sheet'"' != "" local sheetopt sheet(`"`sheet'"')
        quietly import excel `"`filepath'"', `sheetopt' firstrow allstring clear

        // Which columns to process?
        if `"`varname'"' != "" {
            foreach c of local varname {
                capture confirm variable `c', exact
                if _rc {
                    display as error `"Column "`c'" not found in the sheet."'
                    display as error "Note: -import excel, firstrow- converts headers to valid"
                    display as error "Stata names (e.g., spaces become underscores)."
                    exit 111
                }
            }
            local cols `varname'
        }
        else {
            unab cols : _all
        }

        // Check every global name before reading anything
        foreach col of local cols {
            local g `prefix'`col'
            capture confirm name `g'
            if _rc | strlen("`g'") > 32 {
                display as error `"Cannot create global "`g'": not a valid name or longer than 32 characters."'
                exit 198
            }
        }

        // Build each list in a local first
        foreach col of local cols {
            local list_`col' ""
            forvalues i = 1/`=_N' {
                if strpos(`col'[`i'], char(96)) | strpos(`col'[`i'], char(34)) ///
                    | strpos(`col'[`i'], char(36)) {
                    display as error "Column `col', Excel row `=`i'+1': cell contains a"
                    display as error "backtick, double quote, or \$, which cannot go in a macro list."
                    exit 198
                }
                local cell = strtrim(`col'[`i'])
                if `"`cell'"' == "" | `"`cell'"' == "." continue
                local list_`col' `list_`col'' `cell'
            }
            if "`unique'" != "" {
                local list_`col' : list uniq list_`col'
            }
        }

        cwf `caller_frame'

        // All columns read and checked: now assign the globals
        foreach col of local cols {
            local g `prefix'`col'
            global `g' `list_`col''
            local globals `globals' `g'

            if "`verbose'" != "" {
                local n : word count `list_`col''
                display as text `"  \$`g' (`n' items): `list_`col''"'
            }
        }
    }

    local rc = _rc
    capture cwf `caller_frame'
    capture frame drop `tmp_frame'

    if `rc' {
        display as error "load_varlist failed (return code `rc')."
        exit `rc'
    }

    if "`verbose'" != "" display as text "load_varlist: done."

    return local globals `globals'
end
