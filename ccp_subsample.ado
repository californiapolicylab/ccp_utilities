*! ccp_subsample.ado
*! Version 1.0.0
*! Single source of truth for UC-CCP person (ctk) subsamples.
*!
*! Every program that pulls UC-CCP data by subsample (ccp_query, the
*! dashboard's ccd_pull_geo, ...) gets its SQL from here, so the same
*! subsample() value always selects the same people in every table.
*!
*! Syntax:
*!   ccp_subsample pct [, column(sql_expression)]
*!
*!   pct       pt01 pt1 1 2 5 10 20 50 100   (pt01 = 0.01%, pt1 = 0.1%)
*!   column()  SQL expression for the person identifier the condition is
*!             written against; default attribute.ctk. Use the alias of the
*!             table being queried, e.g. column(g.ctk).
*!
*! Returns:
*!   r(where)    SQL condition in parentheses, or "" for the 100% sample
*!   r(pct)      the subsample token as given (lower case)
*!   r(percent)  the subsample as a number, e.g. 2 or 0.01
*!   r(scale)    100 / r(percent): multiplier from sample counts to totals
*!
*! Example:
*!   ccp_subsample 2, column(g.ctk)
*!   di "`r(where)'"
*!   -> (substring(g.ctk, 12, 2) = '64' OR substring(g.ctk, 12, 2) = '69')

cap program drop ccp_subsample
program define ccp_subsample, rclass
    version 16.0
    syntax anything(name=pct) [, COLumn(string)]

    local pct = lower(strtrim(`"`pct'"'))
    if `"`column'"' == "" local column "attribute.ctk"
    local column = strtrim(`"`column'"')
    if !regexm(`"`column'"', "^[A-Za-z_][A-Za-z0-9_.]*$") {
        di as error `"ccp_subsample: column() "`column'" is not a valid SQL column reference."'
        exit 198
    }
    local c `column'

    // Subsample definitions. Edit here only; every caller picks up the change.
    if "`pct'" == "pt01" {
        local w "substring(`c', 12, 2) = '65' AND substring(`c', 9, 2) = '17'"
        local percent 0.01
    }
    else if "`pct'" == "pt1" {
        local w "substring(`c', 12, 2) = '65' AND substring(`c', 9, 1) = '1'"
        local percent 0.1
    }
    else if "`pct'" == "1" {
        local w "substring(`c', 12, 2) = '65'"
        local percent 1
    }
    else if "`pct'" == "2" {
        local w "substring(`c', 12, 2) = '64' OR substring(`c', 12, 2) = '69'"
        local percent 2
    }
    else if "`pct'" == "5" {
        local w "substring(`c', 12, 2) = '17' OR substring(`c', 12, 2) = '32' OR substring(`c', 12, 2) = '85' OR substring(`c', 12, 2) = '90' OR substring(`c', 12, 2) = '56'"
        local percent 5
    }
    else if "`pct'" == "10" {
        local w "substring(`c', 12, 1) = '1'"
        local percent 10
    }
    else if "`pct'" == "20" {
        local w "substring(`c', 12, 1) = '1' OR substring(`c', 12, 1) = '7'"
        local percent 20
    }
    else if "`pct'" == "50" {
        local w "substring(`c', 12, 1) = '1' OR substring(`c', 12, 1) = '7' OR substring(`c', 12, 1) = '2' OR substring(`c', 12, 1) = '5' OR substring(`c', 12, 1) = '8'"
        local percent 50
    }
    else if "`pct'" == "100" {
        local w ""
        local percent 100
    }
    else {
        di as error `"ccp_subsample: "`pct'" is not a valid subsample. Use one of: pt01 pt1 1 2 5 10 20 50 100"'
        exit 198
    }

    if `"`w'"' != "" local w "(`w')"

    return local  where   `"`w'"'
    return local  pct     "`pct'"
    return scalar percent = `percent'
    return scalar scale   = 100 / `percent'
end
