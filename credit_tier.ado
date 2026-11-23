*! credit_tier v1.0.0  23nov2026
*! Classify a credit score (300-850 scale) into credit tiers.
*! Tier cutoffs follow Experian's VantageScore 3.0 ranges.
*!
*! Syntax:
*!   credit_tier scorevar [if] [in] [, GENerate(newvar) LABELname(lblname) REPLACE]
*!
*! Defaults: generate(credit_tier), labelname(<newvar>_lbl)
*! Scores outside 300-850 (and missing scores) are set to missing.

program define credit_tier
    version 14
    syntax varname(numeric) [if] [in] ///
        [, GENerate(name) LABELname(name) REPLACE]

    if "`generate'"  == "" local generate  credit_tier
    if "`labelname'" == "" local labelname `generate'_lbl

    if "`replace'" != "" capture drop `generate'
    else confirm new variable `generate'

    marksample touse

    quietly {
        count if `touse' & !inrange(`varlist', 300, 850)
        local nbad = r(N)

        generate byte `generate' = .
        replace `generate' = 1 if `touse' & `varlist' >= 781 & `varlist' <= 850
        replace `generate' = 2 if `touse' & `varlist' >= 661 & `varlist' <  781
        replace `generate' = 3 if `touse' & `varlist' >= 601 & `varlist' <  661
        replace `generate' = 4 if `touse' & `varlist' >= 501 & `varlist' <  601
        replace `generate' = 5 if `touse' & `varlist' >= 300 & `varlist' <  501
    }

    label define `labelname'                ///
        1 "1=Superprime (781-850)"          ///
        2 "2=Prime (661-780)"               ///
        3 "3=Near prime (601-660)"          ///
        4 "4=Subprime (501-600)"            ///
        5 "5=Deep subprime (300-500)"       ///
        , replace
    label values `generate' `labelname'
    label variable `generate' "Credit tier of `varlist'"

    if `nbad' > 0 {
        display as text "note: `nbad' nonmissing value(s) of `varlist' outside 300-850 set to missing"
    }
end
