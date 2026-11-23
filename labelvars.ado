*! labelvars.ado
*! Version 1.2.0  2026-11-23
*! Apply variable labels from an Excel driver file
*!
*! Syntax:
*!   labelvars using <filepath> [, sheet(sheetname) overwrite]
*!
*! Arguments:
*!   using <filepath>   Path to the Excel driver file (required)
*!   sheet(sheetname)   Excel sheet to read; defaults to the first sheet
*!   overwrite          Overwrite existing variable labels (default: preserve)
*!
*! Driver rules:
*!   - Variable names must match exactly (no abbreviations or wildcards).
*!   - Rows with a blank variable name or a blank label are skipped.
*!   - If a variable appears in more than one row, the first row is used.
*!   - Labels longer than 80 characters are truncated by Stata (flagged).
*!   - Column detection is case-insensitive and follows the priority order
*!     of the candidate lists below, not the column order in the sheet.
*!
*! Stored results:
*!   r(N_labelled)   r(N_skipped)    r(N_notfound)
*!   r(N_blank)      r(N_duplicate)  r(N_truncated)
*!   r(varname_col)  r(label_col)
*!
*! Requires Stata 16 or later (frames).
*!
*! History:
*!   1.2.0  tempname frame (no longer advances the RNG); exact name matching;
*!          priority-based column detection; duplicate, blank-label and
*!          label-length checks; stored results; c_local handoff removed.
*!   1.1.0  2026-04-30

program define labelvars, rclass
    version 16.0

    syntax using/ [, SHEET(string) OVERWRITE]

    confirm file `"`using'"'

    // ------------------------------------------------------------------ //
    // Temporary frame. tempname avoids name collisions without touching
    // the random-number state, and Stata drops tempname frames on exit.
    // ------------------------------------------------------------------ //
    local caller_frame = c(frame)
    tempname tmp_frame
    frame create `tmp_frame'

    local n_applied   = 0
    local n_skipped   = 0
    local n_notfound  = 0
    local n_blank     = 0
    local n_dup       = 0
    local n_truncated = 0
    local varname_col ""
    local label_col   ""

    capture noisily {

        // -------------------------------------------------------------- //
        // Load the driver. We switch frames with cwf (rather than a
        // frame block) so all locals stay in this program's scope.
        // -------------------------------------------------------------- //
        cwf `tmp_frame'

        local sheetopt
        if `"`sheet'"' != "" local sheetopt sheet(`"`sheet'"')
        quietly import excel `"`using'"', `sheetopt' firstrow allstring clear

        unab cols : _all

        // -------------------------------------------------------------- //
        // Detect columns. Outer loop = candidate priority, so e.g. a
        // "varname" column always beats a "name" column wherever it sits.
        // -------------------------------------------------------------- //
        foreach cand in varname var_name variable variable_name var name {
            foreach col of local cols {
                if lower("`col'") == "`cand'" local varname_col `col'
            }
            if "`varname_col'" != "" continue, break
        }

        foreach cand in label varlabel var_label description desc lbl {
            foreach col of local cols {
                if lower("`col'") == "`cand'" local label_col `col'
            }
            if "`label_col'" != "" continue, break
        }

        if "`varname_col'" == "" {
            di as error "labelvars: could not identify a variable-name column."
            di as error "  Expected a column named: varname, var_name, variable, variable_name, var, or name."
            di as error "  Columns found in the driver file: `cols'"
            exit 111
        }

        if "`label_col'" == "" {
            di as error "labelvars: could not identify a label column."
            di as error "  Expected a column named: label, varlabel, var_label, description, desc, or lbl."
            di as error "  Columns found in the driver file: `cols'"
            exit 111
        }

        di as text "labelvars: using column " as result "`varname_col'" ///
            as text " for variable names and " as result "`label_col'"  ///
            as text " for labels."

        // Read driver rows into locals (evaluated directly from the data,
        // so quotes or macro characters in labels are not re-parsed)
        local N_rows = _N
        forvalues i = 1/`N_rows' {
            local vn_`i'  = strtrim(`varname_col'[`i'])
            local lbl_`i' = strtrim(`label_col'[`i'])
        }

        cwf `caller_frame'

        // -------------------------------------------------------------- //
        // Apply the labels in the caller's frame
        // -------------------------------------------------------------- //
        local seen ""

        forvalues i = 1/`N_rows' {
            local vn  `"`vn_`i''"'
            local lbl `"`lbl_`i''"'
            local xlrow = `i' + 1    // Excel row number (row 1 = header)

            // Blank variable name: skip quietly
            if `"`vn'"' == "" {
                local ++n_blank
                continue
            }

            // Must be a single valid name that exists exactly as written
            local ok = 0
            local nwords : word count `vn'
            if `nwords' == 1 {
                capture confirm name `vn'
                if !_rc {
                    capture confirm variable `vn', exact
                    if !_rc local ok = 1
                }
            }
            if !`ok' {
                di as text "  [not found]   `vn'"
                local ++n_notfound
                continue
            }

            // Duplicate rows: first occurrence wins
            if `: list vn in seen' {
                di as text "  [duplicate]   `vn' (Excel row `xlrow'; first occurrence used)"
                local ++n_dup
                continue
            }
            local seen `seen' `vn'

            // Blank label: skip rather than clearing an existing label
            if `"`lbl'"' == "" {
                di as text "  [no label]    `vn' (Excel row `xlrow')"
                local ++n_blank
                continue
            }

            local existing_label : variable label `vn'

            if `"`existing_label'"' != "" & "`overwrite'" == "" {
                di as text "  [skipped]     `vn' (already labelled)"
                local ++n_skipped
                continue
            }

            if ustrlen(`"`lbl'"') > 80 {
                di as text "  [truncated]   `vn' (label is " ///
                    ustrlen(`"`lbl'"') " characters; Stata keeps 80)"
                local ++n_truncated
            }

            label variable `vn' `"`lbl'"'
            if `"`existing_label'"' != "" {
                di as text "  [overwritten] `vn'"
            }
            else {
                di as text "  [labelled]    `vn'"
            }
            local ++n_applied
        }

        // -------------------------------------------------------------- //
        // Summary
        // -------------------------------------------------------------- //
        di as text " "
        di as result "labelvars complete:"
        di as text "  Variables labelled     : " as result `n_applied'
        di as text "  Skipped (no overwrite) : " as result `n_skipped'
        di as text "  Not found in dataset   : " as result `n_notfound'
        di as text "  Duplicate driver rows  : " as result `n_dup'
        di as text "  Blank name or label    : " as result `n_blank'
        if `n_truncated' {
            di as text "  Labels truncated (>80) : " as result `n_truncated'
        }

    } // end capture noisily

    // Always return to the caller's frame and drop the temp frame
    local rc = _rc
    capture cwf `caller_frame'
    capture frame drop `tmp_frame'
    if `rc' exit `rc'

    return scalar N_labelled  = `n_applied'
    return scalar N_skipped   = `n_skipped'
    return scalar N_notfound  = `n_notfound'
    return scalar N_blank     = `n_blank'
    return scalar N_duplicate = `n_dup'
    return scalar N_truncated = `n_truncated'
    return local  varname_col   "`varname_col'"
    return local  label_col     "`label_col'"
end
