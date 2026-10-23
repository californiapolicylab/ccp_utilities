*! backfill_mode v0.1
*! Fills missing values of each variable with the person's modal non-missing
*! value from year from() onward. Ties go to the most recently observed value.
*! Works for numeric and string variables.
*!
*! usage:
*!   backfill_mode birth_date gender_cd
*!   backfill_mode birth_date, id(ctk) year(file_year) from(2011) overwrite
*!
*! options:
*!   id()       person identifier            (default ctk)
*!   year()     panel year variable          (default file_year)
*!   from()     first year used for the mode (default 2011)
*!   overwrite  set every year to the person's mode (makes the variable
*!              constant within person); default fills only missing values
*!              in years before from()
*!
*! returns r(n_changed_<var>) and r(n_missing_<var>) for each variable
*! (names are truncated if the variable name is long)

cap program drop backfill_mode
program define backfill_mode, rclass
	version 16

	syntax varlist [, id(varname) year(varname) from(integer 2011) overwrite]

	if "`id'"   == "" local id   "ctk"
	if "`year'" == "" local year "file_year"
	confirm variable `id' `year'

	local clash : list varlist & id
	local clash `clash' `: list varlist & year'
	if "`clash'" != "" {
		di as err "cannot backfill the id or year variable (`clash')"
		exit 198
	}

	foreach v of local varlist {

		qui count if `year' >= `from' & !missing(`v')
		if r(N) == 0 {
			di as err "`v': no non-missing values from `from' onward; skipped"
			continue
		}

		//modal value per person from years >= from, ties -> most recent
		tempname fr
		tempvar n last neg_n neg_last mode
		frame put `id' `year' `v' if `year' >= `from' & !missing(`v'), into(`fr')
		frame `fr' {
			qui gcollapse (count) `n' = `year' (max) `last' = `year', by(`id' `v')
			qui gen long `neg_n'    = -`n'
			qui gen int  `neg_last' = -`last'
			qui bysort `id' (`neg_n' `neg_last'): keep if _n == 1
			keep `id' `v'
		}

		//bring the mode back
		qui frlink m:1 `id', frame(`fr')
		qui frget `mode' = `v', from(`fr')

		//fill
		if "`overwrite'" != "" {
			qui count if `v' != `mode' & !missing(`mode')
			local nchg = r(N)
			qui replace `v' = `mode' if !missing(`mode')
			di as txt "`v': " %12.0fc `nchg' " values replaced with person mode"
		}
		else {
			qui count if missing(`v') & `year' < `from' & !missing(`mode')
			local nchg = r(N)
			qui replace `v' = `mode' if missing(`v') & `year' < `from'
			di as txt "`v': " %12.0fc `nchg' " pre-`from' missing values filled"
		}

		qui count if missing(`v')
		local nmiss = r(N)
		di as txt "`v': " %12.0fc `nmiss' " values still missing"

		drop `fr' `mode'
		frame drop `fr'

		local rn = substr("`v'", 1, 20)
		return scalar n_changed_`rn' = `nchg'
		return scalar n_missing_`rn' = `nmiss'
	}
end
