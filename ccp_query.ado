*! ccp_query.ado
*! Version 1.5.0
*! Executes ODBC SQL queries across quarterly data archives
*! Archives are named by scope (ca/national), quarter-month (MMYYYY), and source
*! (attribute, trade, inquiry, public)
*!
*! CHANGELOG (1.1.0):
*!   NEW: extra_tables()/extra_keys()/extra_aliases() options add support
*!        infrastructure for LEFT JOINing in one or more additional tables
*!        that are not among the four hard-coded sources (attribute/trade/
*!        inquiry/public), each on its own (possibly multi-field) join key.
*!        See section 2b below. No changes were made to the basic_cleaning
*!        section as part of this update.
*!   FIX:  national archive tables are now queried as national_MMYYYY_<source>
*!        (previously ntl_MMYYYY_<source>). sample(ntl) is unchanged, and
*!        _archive_scope / output file names still use "ntl".
*!
*! CHANGELOG (1.2.0):
*!   FIX:  sql_where_*/sql_if_*/sql_on_* macros renamed to full source
*!        names so their lookups actually resolve, and table aliases in them
*!        corrected (att./tr. -> attribute./trade.). The attribute WHERE
*!        filters (PR source, not deceased, Vantage > 300) are now applied
*!        in SQL; previously they were silently skipped.
*!   NEW:  optional "expression" column in the Excel driver file for derived
*!        fields (see section 3). Derived fields are computed in SQL and can
*!        be selected and used as join keys like ordinary columns.
*!   FIX:  extra_aliases() now defaults to the last dotted part of the table
*!        name (e.g. addr_geo_xwalk), since a full schema-qualified name is
*!        not a valid Stata macro name. Aliases are validated.
*!   No changes were made to the basic_cleaning section.
*!
*! CHANGELOG (1.3.0):
*!   NEW:  parquet option saves each quarterly archive file as Parquet
*!        (via the CPL pqsave/pquse/pqtools tools) instead of .dta. The
*!        combined output named by name_output() is still a .dta.
*!        pqsave_options() passes extra options (e.g. compression(snappy))
*!        through to pqsave. Prerequisites are checked before any queries run
*!        (section 5b). No changes were made to the basic_cleaning section.
*!
*! CHANGELOG (1.4.0):
*!   NEW:  filter_file()/filter_table()/filter_mode() options restrict a
*!        pull to people associated with a user-supplied list of values
*!        (e.g. school IDs, counties, ZIPs) in a multi-period crosswalk
*!        table, default constructed_data.addr_geo_xwalk. See section 2c.
*!        The filter is ANDed with the existing WHERE filters and
*!        subsample(). Requires attribute variables in the pull.
*!        filter_mode(archive) matches the filter table's year/month columns
*!        to each attribute record's file_date. Filter tables must be
*!        multi-period tables (like the xwalk), not per-archive UC-CCP
*!        tables. No changes were made to the basic_cleaning section.
*!
*! CHANGELOG (1.5.0):
*!   CHANGED (basic_cleaning): geography cleaning in phase 2 is now
*!        source-aware. Geography from the attribute tables is renamed and
*!        checked as before; geography from a geo crosswalk joined through
*!        extra_tables() (any table named addr_geo_xwalk*, e.g.
*!        addr_geo_xwalk_v2) is renamed to matching geo_* names and left
*!        otherwise untouched, since it is validated when the crosswalk is
*!        built. If both are pulled, the attribute versions get an _attr
*!        suffix. See section 3b. No other basic_cleaning steps changed.
*!   FIX (basic_cleaning): the step that blanks non-numeric geography codes
*!        compared each variable's first value, not its name, to
*!        "geo_stabbr". It blanked every state abbreviation and stopped with
*!        a type-mismatch error if any geo_* variable was numeric.
*!   NEW:  an extra table's column whose name is already in the pull (e.g.
*!        zip_cd from both attribute and the crosswalk) is now selected as
*!        <column>_<alias> instead of producing a duplicate column, which
*!        made the ODBC load fail.
*!
*! Syntax:
*!   ccp_query, sample(ca|ntl|both) dir_varlist(path) name_output(name)
*!                   dir_output(path) archives(spec)
*!                   [append_existing(path) subsample(pct) basic_cleaning
*!                   extra_tables(tbl1 tbl2 ...) extra_keys(k1 k2 | k3 ...)
*!                   extra_aliases(alias1 alias2 ...)
*!                   parquet pqsave_options(string)
*!                   filter_file(path.csv) filter_table(tbl)
*!                   filter_mode(ever|archive)]
cap program drop ccp_query
program define ccp_query
    version 16.0

    syntax , ///
        SAMple(string)          /// ca, ntl, or both (also accepts "california", "national", or capitalized versions of those)
        DIR_varlist(string)     /// path to Excel file with variable list
        NAME_output(string)     /// name for output .dta file (no extension needed)
        DIR_output(string)      /// directory for output file
        ARCHives(string)        /// archive range/list specification
        [                       ///
        APPEND_existing(string) /// path to existing .dta file to append results into
        SUBsample(string)       /// subsample size: pt01 pt1 1 2 5 10 20 50 100
        BASIC_cleaning          /// apply basic cleaning steps before appending archives
		SHEET(string)			/// sheet of the Excel driver file to use
        EXTRA_tables(string)    /// NEW: space-separated list of additional (non-hard-coded) table names to LEFT JOIN
        EXTRA_keys(string)      /// NEW: join key(s) for extra_tables(); one space-separated block per table, blocks separated by "|"
        EXTRA_aliases(string)   /// NEW: optional SQL aliases for extra_tables(), space-separated, parallel list (defaults to table name)
        PARQuet                 /// NEW (1.3.0): save quarterly archive files as .parquet instead of .dta
        PQSAVE_options(string)  /// NEW (1.3.0): extra options passed to pqsave (not name(), directory(), or replace)
        filter_file(string)     /// NEW (1.4.0): CSV with one column; header = filter-table column name, rows = values to keep
        filter_table(string)    /// NEW (1.4.0): table the filter values are looked up in (default constructed_data.addr_geo_xwalk)
        filter_mode(string)     /// NEW (1.4.0): ever (default) | archive
        ]

    // =========================================================================
    // 0.  PROGRAMMER-DEFINED MACROS
    //     Fill these in before deploying. They are intentionally left as
    //     empty (or placeholder) locals so nothing breaks at parse time.
    // =========================================================================

    // --- IF clause (appended to each per-source SELECT, e.g. "IF col = val") ---
    // FIX (1.2.0): these macros are looked up later as sql_if_`src', where
    //   `src' is the full source name (attribute/trade/inquiry/public). They
    //   were previously named with abbreviations (attr/inq/pub), so the
    //   lookups silently returned "". Renamed to the full source names.
    local sql_if_attribute ""   /* attribute source  */
    local sql_if_trade     ""   /* trade source      */
    local sql_if_inquiry   ""   /* inquiry source    */
    local sql_if_public    ""   /* public source     */

    // --- WHERE clause fragments ----------
    // FIX (1.2.0): renamed to full source names (see note above) and table
    //   alias changed from att. to attribute., which is the alias actually
    //   used in the FROM clause. This WHERE clause was previously never
    //   applied; it now is, whenever attribute variables are requested.
    local sql_where_attribute "attribute.source_cd_pr='PR' AND isnull(attribute.deceased_in,'N') <> 'Y' AND attribute.vantage_v4_score IS NOT NULL AND attribute.vantage_v4_score > 300"
    local sql_where_trade     ""
    local sql_where_inquiry   ""
    local sql_where_public    ""

    // --- ON clause for each LEFT JOIN (e.g. "attr.id = trade.id") -------------
    // FIX (1.2.0): renamed to full source names and aliases (att./tr. ->
    //   attribute./trade.) so the sql_on_`primary_src'_`src' lookup finds
    //   them. The attribute-trade join is unchanged in effect (it previously
    //   fell back to join_key = record_nb, which is the same condition).
    local sql_on_attribute_trade   "attribute.record_nb = trade.record_nb"   /* join attribute → trade   */
    local sql_on_attribute_inquiry ""   /* join attribute → inquiry */
    local sql_on_attribute_public  ""   /* join attribute → public  */

	// TODO: this chunk needs to change to acknowledge that there are 
	//			  multiple join keys (record nb and ctk); for now, just using 
	//			  attribute-tradeline case, which is all most projects need
	// 
	
    // --- Primary key / join key used in LEFT JOINs ----------------------------
    //     Set this to whatever unique row identifier the archives share
    local join_key  "record_nb"  

    // --- NEW (1.4.0): defaults for filter_file() (see section 2c) -------------
    //     filter_table_default: table searched when filter_table() is omitted
    //     filter_ctk_col:       person identifier in the filter table
    //     filter_year_col / filter_month_col: period columns in the filter
    //       table, used only by filter_mode(archive). They are compared to
    //       YEAR() and MONTH() of filter_date_expr, the date of each
    //       attribute record. This works whether the columns are stored as
    //       integers or as numeric strings such as '03'.
    local filter_table_default "constructed_data.addr_geo_xwalk"
    local filter_ctk_col       "ctk"
    local filter_year_col      "file_year"
    local filter_month_col     "file_month"
    local filter_date_expr     "attribute.file_date"   /* date type in the attribute tables */
    local filter_alias         "ccp_filter"   /* reserved SQL alias for the filter subquery */
    local filter_warn_n        5000           /* warn above this many filter values */

    // --- NEW (1.5.0): geography sources for basic_cleaning (see section 3b) ---
    //     geo_xwalk_pattern: an extra table whose name (after the schema)
    //       matches this is treated as a geo crosswalk
    //     geo_attr_map: attribute geography column -> cleaned name, in pairs
    //     geo_xw_map / geo_xw_vint_map: crosswalk column -> cleaned name; the
    //       vintage map handles <prefix>_YYYY columns. Crosswalk columns not
    //       listed (school zones, addr_id, flags, ...) become geo_<column>.
    local geo_xwalk_pattern "^addr_geo_xwalk"
    local geo_attr_map    "zip_cd geo_zip state_cd geo_stabbr fips_state geo_state fips_county geo_county census_tract geo_tract census_blk_grp geo_blkgrp census_blk_id geo_blk"
    local geo_xw_map      "zip_cd geo_zip zip_cd_reported geo_zip_reported state_cd geo_stabbr fips_state_cd geo_state county_fips3 geo_county"
    local geo_xw_vint_map "census_tract_cd geo_tract census_blk_grp_cd geo_blkgrp census_blk_id geo_blk census_blk_full_cd geo_blkfull census_tract_geoid geo_tractid"

    // --- ODBC data-source name (DSN) ------------------------------------------
    local dsn  "UCCCP-Asteria"   

    // --- NEW (1.3.0): where users can get the Parquet tools (pqsave/pquse/
    //     pqtools.py). Fill in the repo URL on CPL's OneDev instance; if left blank, the
    //     install message just points users to OneDev generically.
    local cpl_pq_onedev_url ""

    // --- SQL table aliases (fixed; do not change) -----------------------------
    local alias_attribute "att"
    local alias_trade     "tr"
    local alias_inquiry   "inq"
    local alias_public    "pub"

    // --- Subsample WHERE fragments ------------------
    //     Each macro should be a valid SQL WHERE-clause fragment.
    local subsample_pt01  "substring(attribute.ctk, 12, 2) = '65' AND substring(attribute.ctk, 9, 2) = '17'"   /* 0.01 % sample */
    local subsample_pt1   "substring(attribute.ctk, 12, 2) = '65' AND substring(attribute.ctk, 9, 1) = '1'"   /* 0.1  % sample */
    local subsample_1     "substring(attribute.ctk, 12, 2) = '65'"   /* 1    % sample */
    local subsample_2     "substring(attribute.ctk, 12, 2) = '64' OR substring(attribute.ctk, 12, 2) = '69'"   /* 2    % sample */
    local subsample_5     "(substring(attribute.ctk, 12, 2) = '17' OR substring(attribute.ctk, 12, 2) = '32' OR substring(attribute.ctk, 12, 2) = '85' OR substring(attribute.ctk, 12, 2) = '90' OR substring(attribute.ctk, 12, 2) = '56')"   /* 5    % sample */
    local subsample_10    "substring(attribute.ctk, 12, 1) = '1'"   /* 10   % sample */
    local subsample_20    "(substring(attribute.ctk, 12, 1) = '1' OR substring(attribute.ctk, 12, 1) = '7')"   /* 20   % sample */
    local subsample_50    "(substring(attribute.ctk, 12, 1) = '1' OR substring(attribute.ctk, 12, 1) = '7' OR substring(attribute.ctk, 12, 1) = '2' OR substring(attribute.ctk, 12, 1) = '5' OR substring(attribute.ctk, 12, 1) = '8')"   /* 50   % sample */
    local subsample_100   ""   /* 100  % (full) sample */

    // =========================================================================
    // 1.  VALIDATE AND NORMALISE sample() ARGUMENT
    // =========================================================================

    local sample_lc = lower(trim("`sample'"))

    if inlist("`sample_lc'", "ca", "california") {
        local scopes "ca"
    }
    else if inlist("`sample_lc'", "ntl", "national", "both", "ca ntl", "ntl ca") {
        if "`sample_lc'" == "both" | "`sample_lc'" == "ca ntl" | "`sample_lc'" == "ntl ca" {
            local scopes "ca ntl"
        }
        else {
            local scopes "ntl"
        }
    }
    else {
        di as error "sample() must be ca, ntl, or both (also accepts 'california', 'national')."
        exit 198
    }

    // =========================================================================
    // 2.  VALIDATE subsample() ARGUMENT (if supplied)
    // =========================================================================

	* If subsample is provided, check that it makes sense and define SQL
	*	code appropriately
    if `"`subsample'"' != "" {
        local sub_lc = lower(trim(`"`subsample'"'))
        local valid_subs "pt01 pt1 1 2 5 10 20 50 100"
        local sub_ok 0
        foreach v of local valid_subs {
            if "`sub_lc'" == "`v'" local sub_ok 1
        }
        if `sub_ok' == 0 {
            di as error "subsample() must be one of: pt01 pt1 1 2 5 10 20 50 100"
            exit 198
        }
        // Map subsample value to the appropriate programmer-defined macro
        local subsample_where "`subsample_`sub_lc''"
    }

	* If subsample isn't provided, default to the 0.01% sample (smallest) and 
	*	 inform the user
	if mi(`"`subsample'"') {
		local upper_sample = upper("`sample'")
		di ""
		di as text "NOTE: We noticed that you did not supply a subsample. The UC-CCP is very large and it is unwise to pull down a 100% sample."
		di as text "	By default, this program produces a random 0.01% sample of the `upper_sample' database(s)."
		if `"`filter_file'"' != "" {
			di as text "	filter_file() is applied within that 0.01% sample. Use subsample(100) to get everyone matching the filter."
		}
		local subsample "pt01"
		local subsample_where "`subsample_pt01'"
	}
	
    // =========================================================================
    // 2b. NEW: PARSE EXTRA JOIN TABLE SPECIFICATION (flexible join support)
    //     Lets the user LEFT JOIN in one or more additional tables that are
    //     NOT among the four hard-coded sources (attribute/trade/inquiry/
    //     public). Each extra table is joined against the primary source
    //     table on one or more join keys (fields assumed to share the same
    //     name in both tables, mirroring how join_key works for the four
    //     built-in sources above).
    //
    //       extra_tables("tbl1 tbl2")       space-separated literal SQL
    //                                       table names, used verbatim (not
    //                                       archive/scope-prefixed)
    //       extra_keys("k1 k2 | k3")        one join-key block per table in
    //                                       extra_tables(), blocks separated
    //                                       by "|"; each block is a
    //                                       space-separated list of fields
    //       extra_aliases("a1 a2")          optional, parallel to
    //                                       extra_tables(); defaults to the
    //                                       table name itself
    //
    //     This section only stores the parsed specification (extra_tblname_`k',
    //     extra_alias_`k', extra_keys_`k', extra_alias_list, n_extra_tables).
    //     It is consumed later: extra_alias_list is added to the set of
    //     valid Excel "source" values (section 3) and to the SELECT column
    //     list (section 6a), and the LEFT JOIN clauses themselves are built
    //     in section 6a from extra_tblname_`k'/extra_alias_`k'/extra_keys_`k'.
    // =========================================================================

    local n_extra_tables 0
    local extra_alias_list ""

    if `"`extra_tables'"' != "" {

        local extra_tbl_list = trim(`"`extra_tables'"')
        local n_extra_tables : word count `extra_tbl_list'

        if `"`extra_keys'"' == "" {
            di as error "extra_keys() must be supplied when extra_tables() is specified."
            exit 198
        }

        // Split extra_keys() on "|" into one key-block per table. Each block
        // is stored directly under raw_keyblock_`n_key_blocks' (rather than
        // re-concatenated and re-split by whitespace), since a block itself
        // may contain multiple space-separated fields for a composite key.
        local keys_remaining = trim(`"`extra_keys'"')
        local n_key_blocks 0
        while `"`keys_remaining'"' != "" {
            local bar_pos = strpos(`"`keys_remaining'"', "|")
            if `bar_pos' == 0 {
                local this_block = trim(`"`keys_remaining'"')
                local keys_remaining ""
            }
            else {
                local this_block = trim(substr(`"`keys_remaining'"', 1, `bar_pos' - 1))
                local keys_remaining = trim(substr(`"`keys_remaining'"', `bar_pos' + 1, .))
            }
            local ++n_key_blocks
            local raw_keyblock_`n_key_blocks' "`this_block'"
        }

        if `n_key_blocks' != `n_extra_tables' {
            di as error "extra_tables() lists `n_extra_tables' table(s) but extra_keys() supplies `n_key_blocks' join-key block(s) (blocks are separated by '|'). These counts must match."
            exit 198
        }

        // Optional aliases. CHANGED (1.2.0): default is now the last dotted
        // part of the table name (constructed_data.addr_geo_xwalk ->
        // addr_geo_xwalk), because the alias is also used inside Stata macro
        // names (varlist_<alias>) and a dotted name is not valid there.
        local extra_alias_list ""
        if `"`extra_aliases'"' != "" {
            local extra_alias_list = lower(trim(`"`extra_aliases'"'))
            local n_aliases : word count `extra_alias_list'
            if `n_aliases' != `n_extra_tables' {
                di as error "extra_aliases() supplies `n_aliases' alias(es) but extra_tables() lists `n_extra_tables' table(s). These counts must match."
                exit 198
            }
        }
        else {
            foreach t of local extra_tbl_list {
                local t_last = lower(substr("`t'", strrpos("`t'", ".") + 1, .))
                local extra_alias_list "`extra_alias_list' `t_last'"
            }
            local extra_alias_list = trim("`extra_alias_list'")
        }

        // Validate aliases: SQL/Stata-safe identifiers, short enough to fit
        // in Stata macro names (31-char limit, less the 8-char "varlist_"
        // prefix), not colliding with built-in sources or each other
        local seen_aliases "attribute trade inquiry public `filter_alias'"   /* CHANGED (1.4.0): filter alias reserved */
        foreach a of local extra_alias_list {
            if !regexm("`a'", "^[a-z_][a-z0-9_]*$") | length("`a'") > 23 {
                di as error `"extra table alias "`a'" is invalid: use letters, digits, and underscores only (max 23 characters). Supply extra_aliases() to set one explicitly."'
                exit 198
            }
            if `: list a in seen_aliases' {
                di as error `"extra table alias "`a'" is duplicated or collides with a built-in source name."'
                exit 198
            }
            local seen_aliases "`seen_aliases' `a'"
        }

        // Store one set of macros per extra table: name, alias, and its
        // (possibly multi-field) join-key list
        forvalues k = 1/`n_extra_tables' {
            local extra_tblname_`k' : word `k' of `extra_tbl_list'
            local extra_alias_`k'   : word `k' of `extra_alias_list'

            local extra_keys_`k' "`raw_keyblock_`k''"
            if "`extra_keys_`k''" == "" {
                di as error "extra_keys(): join-key block `k' (for table `extra_tblname_`k'') is empty."
                exit 198
            }
        }
    }

    // =========================================================================
    // 2c. NEW (1.4.0): PARSE filter_file() / filter_table() / filter_mode()
    //     Restricts the pull to people (ctks) associated with any of a list
    //     of values in one column of a crosswalk table.
    //
    //       filter_file(path.csv)  CSV with exactly one column. The header
    //                              is the filter-table column to match on
    //                              (e.g. primary_school_id, fips_cd); each
    //                              row is one value to keep. Values are read
    //                              as text, so leading zeros in the file are
    //                              preserved (but not restored if lost).
    //       filter_table(tbl)      table to look values up in; default
    //                              `filter_table_default'. Must contain
    //                              the ctk column and the filter column.
    //       filter_mode(mode)      ever    - keep a person in every archive
    //                                        if they EVER have a matching row
    //                                        in the filter table (any period)
    //                              archive - keep an attribute record only if
    //                                        the person's filter-table row for
    //                                        the year and month of that
    //                                        record's file_date matches
    //
    //     The filter table must be a multi-period table such as the geo
    //     xwalk. Filtering on fields from the per-archive UC-CCP tables
    //     (ca_MMYYYY_attribute etc.) is not supported: an "ever" filter
    //     would have to scan every archive, and a within-archive filter on
    //     those fields is better applied after download.
    //
    //     Implemented as one EXISTS condition ANDed into each archive's
    //     WHERE clause (section 6a):
    //       EXISTS (SELECT 1 FROM <tbl> AS ccp_filter
    //               WHERE ccp_filter.ctk = attribute.ctk
    //                 AND ccp_filter.<col> IN ('v1','v2',...)
    //                [AND ccp_filter.file_year  = YEAR(attribute.file_date)
    //                 AND ccp_filter.file_month = MONTH(attribute.file_date)])
    //     For speed, the filter table should have an index on <col> that
    //     includes ctk.
    // =========================================================================

    local has_filter 0
    if `"`filter_file'"' != "" {
        local has_filter 1

        // --- mode ---
        local filter_mode = lower(trim(`"`filter_mode'"'))
        if "`filter_mode'" == "" local filter_mode "ever"
        if !inlist("`filter_mode'", "ever", "archive") {
            di as error "filter_mode() must be ever or archive."
            exit 198
        }

        // --- table ---
        local filter_tbl = trim(`"`filter_table'"')
        if `"`filter_tbl'"' == "" local filter_tbl "`filter_table_default'"
        if !regexm(`"`filter_tbl'"', "^[A-Za-z0-9_.]+$") {
            di as error `"filter_table() "`filter_tbl'" is invalid: use letters, digits, underscores, and dots only (e.g. schema.table)."'
            exit 198
        }
        local filter_tbl_last = lower(substr(`"`filter_tbl'"', strrpos(`"`filter_tbl'"', ".") + 1, .))
        if regexm("`filter_tbl_last'", "^(ca|ntl|national)_[0-9][0-9][0-9][0-9][0-9][0-9]_") {
            di as error `"filter_table() "`filter_tbl'" looks like a per-archive UC-CCP table. Filters must use a multi-period table such as `filter_table_default'."'
            exit 198
        }

        // --- file ---
        capture confirm file `"`filter_file'"'
        if _rc {
            di as error `"filter_file(): file "`filter_file'" not found."'
            exit 601
        }

        preserve
            quietly import delimited using `"`filter_file'"', ///
                varnames(1) stringcols(_all) case(lower) clear

            quietly ds
            local filter_col "`r(varlist)'"
            local n_fcols : word count `filter_col'
            if `n_fcols' != 1 {
                di as error "filter_file() must contain exactly one column (found `n_fcols': `filter_col')."
                di as error "The header should be the name of the filter-table column to match on, e.g. primary_school_id."
                exit 198
            }
            // import delimited renames headers that aren't valid Stata names
            // to v1 etc.; the original header is kept as the variable label
            local filter_hdr : variable label `filter_col'
            if "`filter_col'" == "v1" & `"`filter_hdr'"' != "" & lower(`"`filter_hdr'"') != "v1" {
                di as error `"filter_file(): header "`filter_hdr'" is not a valid column name. Use letters, digits, and underscores only."'
                exit 198
            }
            if !regexm("`filter_col'", "^[a-z_][a-z0-9_]*$") {
                di as error `"filter_file(): header "`filter_col'" is not a valid column name."'
                exit 198
            }

            // Clean values: trim, drop blanks, de-duplicate
            quietly replace `filter_col' = strtrim(`filter_col')
            quietly drop if `filter_col' == ""
            quietly duplicates drop `filter_col', force
            quietly count
            local n_filter_vals = r(N)
            if `n_filter_vals' == 0 {
                di as error "filter_file() contains no non-blank values."
                exit 198
            }

            // Reject characters that would break Stata macros or the SQL string
            quietly count if strpos(`filter_col', char(34)) | strpos(`filter_col', char(96)) | ///
                             strpos(`filter_col', char(36)) | strpos(`filter_col', char(92))
            if r(N) > 0 {
                di as error "filter_file(): `r(N)' value(s) contain a double quote, backtick, dollar sign, or backslash, which are not allowed."
                exit 198
            }

            // Build the quoted IN list, doubling any apostrophes for SQL
            sort `filter_col'
            local filter_in ""
            local filter_preview ""
            forvalues i = 1/`n_filter_vals' {
                local fv = subinstr(`filter_col'[`i'], "'", "''", .)
                if `i' == 1 local filter_in "'`fv''"
                else        local filter_in "`filter_in','`fv''"
                if `i' <= 5 local filter_preview "`filter_preview' `fv'"
            }
        restore

        // Full EXISTS condition. It is the same for every archive: in archive
        // mode the period match is correlated with each attribute record's
        // own file_date rather than taken from the archive name
        local filter_cond "`filter_alias'.`filter_ctk_col' = attribute.ctk AND `filter_alias'.`filter_col' IN (`filter_in')"
        if "`filter_mode'" == "archive" {
            local filter_cond "`filter_cond' AND `filter_alias'.`filter_year_col' = YEAR(`filter_date_expr') AND `filter_alias'.`filter_month_col' = MONTH(`filter_date_expr')"
        }

        di as text _n "Filter applied (filter_file):"
        di as text "  table:  `filter_tbl'"
        di as text "  column: `filter_col'"
        di as text "  values: `n_filter_vals' (first few:`filter_preview')"
        di as text "  mode:   `filter_mode'"
        if `n_filter_vals' > `filter_warn_n' {
            di as text "  WARNING: more than `filter_warn_n' values. Long IN lists can be slow for SQL Server to compile, and may fail."
            di as text "           Consider filtering on a coarser column if one is available."
        }
    }
    else if `"`filter_table'`filter_mode'"' != "" {
        di as error "filter_table() and filter_mode() require filter_file()."
        exit 198
    }

    // =========================================================================
    // 3.  READ VARIABLE LIST FROM EXCEL
    // =========================================================================

    // Import the Excel file. We expect at minimum two columns:
    //   Column 1: variable name
    //   Column 2: source (attribute | trade | inquiry | public)
    // NEW (1.2.0): optional third column "expression". If filled in for a
    //   row, that row's variable is a DERIVED field: an output column named
    //   `variable', computed by the SQL expression against the row's
    //   `source' table, e.g. variable=file_year, source=attribute,
    //   expression=YEAR(file_date). Expressions reference bare column names
    //   (no table alias) and must be valid SQL Server syntax. Avoid
    //   backticks (`) and dollar signs ($), which Stata treats as macro
    //   characters. Leave blank for ordinary pass-through columns.
    preserve
	
		if !mi("`sheet'") local sheet_cmd `"sheet("`sheet'")"'
        quietly import excel using `"`dir_varlist'"', firstrow clear case(lower) `sheet_cmd'

        // Identify the variable-name column and the source column
        // We look for columns literally named "variable" (or "var") and "source"
        capture confirm variable variable
        if _rc {
            capture confirm variable var
            if _rc {
                di as error "Excel file must contain a column named 'variable' or 'var'."
                exit 198
            }
            rename var variable
        }

        capture confirm variable source
        if _rc {
            di as error "Excel file must contain a column named 'source'."
            exit 198
        }

        // Validate source values
		local incl_tradeline = 0
        local valid_sources "attribute trade inquiry public `extra_alias_list'"   /* NEW: extra_alias_list appended for flexible join support */
        quietly glevelsof source, local(src_vals) clean
        foreach sv of local src_vals {
            local sv_lc = lower(trim("`sv'"))
				if "`sv_lc'" == "trade" local incl_tradeline = 1
            local src_ok 0
            foreach vs of local valid_sources {
                if "`sv_lc'" == "`vs'" local src_ok 1
            }
            if `src_ok' == 0 {
                di as error `"Invalid source value "`sv'" found in Excel. Must be one of: `valid_sources'."'
                exit 198
            }
        }
		
        // Build per-source variable lists (NEW: `extra_alias_list' appended for flexible join support)
        foreach src in attribute trade inquiry public `extra_alias_list' {
            local varlist_`src' ""
        }

        // NEW (1.2.0): detect optional expression column. An all-blank
        // column may import as numeric; treat that as "no expressions".
        local has_expr 0
        capture confirm variable expression
        if !_rc {
            capture confirm string variable expression
            if !_rc {
                local has_expr 1
            }
            else {
                quietly count if !missing(expression)
                if r(N) > 0 {
                    di as error "Excel column 'expression' must contain text (SQL expressions)."
                    exit 198
                }
            }
        }
        local n_derived 0

        quietly count
        local nrows = r(N)
        forvalues i = 1/`nrows' {
            local vname = lower(trim(variable[`i']))
            local vsrc  = lower(trim(source[`i']))
            local varlist_`vsrc' `"`varlist_`vsrc'' `vname'"'

            // NEW (1.2.0): record derived fields (stored by index, since
            // source+name combinations can exceed Stata's macro-name limit)
            if `has_expr' {
                local vexpr = trim(expression[`i'])
                if `"`vexpr'"' != "" {
                    if !regexm("`vname'", "^[a-z_][a-z0-9_]*$") | length("`vname'") > 32 {
                        di as error `"Derived field name "`vname'" is invalid: use letters, digits, and underscores only (max 32 characters)."'
                        exit 198
                    }
                    local ++n_derived
                    local derived_src_`n_derived'  "`vsrc'"
                    local derived_name_`n_derived' "`vname'"
                    local derived_expr_`n_derived' `"`vexpr'"'
                }
            }
        }

        // Determine which sources are actually needed (NEW: `extra_alias_list' appended for flexible join support)
        local active_sources ""
        foreach src in attribute trade inquiry public `extra_alias_list' {
            local varlist_`src' = trim(`"`varlist_`src''"')
            if `"`varlist_`src''"' != "" {
                local active_sources "`active_sources' `src'"
            }
        }
        local active_sources = trim("`active_sources'")
    restore

    if "`active_sources'" == "" {
        di as error "No variables found in the Excel variable list."
        exit 198
    }

    // NEW (1.4.0): the filter matches on attribute.ctk, so the pull must
    // include the attribute source (which is then the primary source)
    if `has_filter' & `"`varlist_attribute'"' == "" {
        di as error "filter_file() requires at least one attribute variable in the driver file (e.g. ctk with source = attribute)."
        exit 198
    }

    // NEW (1.2.0): build, for each source, the extra SELECT fragment that
    // computes its derived fields (", <expr> AS <name>, ..."). In section 6a
    // any source with a non-empty fragment is wrapped in a derived table:
    //   (SELECT * , <expr> AS <name> FROM <table>) AS <source>
    // so its derived fields can be selected, joined on, and filtered on by
    // name exactly like ordinary columns.
    foreach src in attribute trade inquiry public `extra_alias_list' {
        local dsel_`src' ""
    }
    forvalues d = 1/`n_derived' {
        local ds "`derived_src_`d''"
        local dsel_`ds' `"`dsel_`ds'', `derived_expr_`d'' AS `derived_name_`d''"'
    }
    if `n_derived' > 0 {
        di as text "Derived fields (computed in SQL):"
        forvalues d = 1/`n_derived' {
            di as text `"  `derived_src_`d''.`derived_name_`d'' = `derived_expr_`d''"'
        }
    }

    // =========================================================================
    // 3b. NEW (1.5.0): SELECT LIST, COLUMN-NAME CLASHES, GEOGRAPHY SOURCES
    //     Built once here (it is the same for every archive) and used in
    //     section 6a (SELECT list) and in basic_cleaning phase 2 (geography).
    //
    //     Clashes: if an extra table's column name is already in the pull,
    //       it is selected as <column>_<alias> so the ODBC load doesn't get
    //       two columns with the same name.
    //     Geography: records which geography columns come from the
    //       attribute tables and which from a geo crosswalk, and the name
    //       each gets in basic_cleaning:
    //         attribute only  -> geo_zip, geo_stabbr, geo_state, geo_county,
    //                            geo_tract, geo_blkgrp, geo_blk (as before)
    //         crosswalk only  -> geo_zip, geo_stabbr, geo_state, geo_county,
    //                            geo_tract_YYYY, geo_blkgrp_YYYY, geo_blk_YYYY,
    //                            geo_blkfull_YYYY, geo_tractid_YYYY, and
    //                            geo_<column> for everything else
    //         both            -> crosswalk as above; attribute columns get
    //                            an _attr suffix (geo_zip_attr, ...)
    //       Every crosswalk column gets a geo_ name so it is kept as a by()
    //       variable if phase 3 collapses tradelines to person-quarters.
    // =========================================================================

    // Which extra tables are geo crosswalks?
    local geo_xw_aliases ""
    forvalues k = 1/`n_extra_tables' {
        local t_last = lower(substr("`extra_tblname_`k''", strrpos("`extra_tblname_`k''", ".") + 1, .))
        if regexm("`t_last'", "`geo_xwalk_pattern'") {
            local geo_xw_aliases "`geo_xw_aliases' `extra_alias_`k''"
        }
    }

    local select_cols_all ""
    local selected_names  ""
    local n_geo_xw 0
    foreach src in attribute trade inquiry public `extra_alias_list' {
        local is_extra  : list src in extra_alias_list
        local is_geo_xw : list src in geo_xw_aliases
        foreach vv of local varlist_`src' {
            local recv "`vv'"
            local clash : list vv in selected_names
            if `is_extra' & `clash' {
                local recv "`vv'_`src'"
                if length("`recv'") > 32 {
                    di as error "Column `vv' from extra table `src' is already in the pull, and `recv' is too long to use instead. Use a shorter alias in extra_aliases()."
                    exit 198
                }
                local select_cols_all "`select_cols_all', `src'.`vv' AS `recv'"
                di as text "NOTE: `src'.`vv' is selected as `recv' (`vv' is already in the pull)."
            }
            else {
                local select_cols_all "`select_cols_all', `src'.`vv'"
            }
            local selected_names "`selected_names' `recv'"

            // Cleaned name for each crosswalk column
            if `is_geo_xw' {
                local t ""
                local nm : word count `geo_xw_map'
                forvalues w = 1(2)`nm' {
                    if "`vv'" == "`: word `w' of `geo_xw_map''" {
                        local t : word `=`w'+1' of `geo_xw_map'
                    }
                }
                local nm : word count `geo_xw_vint_map'
                forvalues w = 1(2)`nm' {
                    local pfx : word `w' of `geo_xw_vint_map'
                    if "`t'" == "" & regexm("`vv'", "^`pfx'_([0-9][0-9][0-9][0-9])$") {
                        local t = "`: word `=`w'+1' of `geo_xw_vint_map''_" + regexs(1)
                    }
                }
                if "`t'" == "" {
                    if substr("`vv'", 1, 4) == "geo_" local t "`vv'"
                    else                               local t "geo_`vv'"
                }
                if length("`t'") > 32 local t "`recv'"
                local ++n_geo_xw
                local geo_xw_recv_`n_geo_xw' "`recv'"
                local geo_xw_tgt_`n_geo_xw'  "`t'"
            }
        }
    }

    // Attribute geography columns in the pull, and their cleaned names
    local n_geo_attr 0
    local nm : word count `geo_attr_map'
    forvalues w = 1(2)`nm' {
        local o : word `w' of `geo_attr_map'
        if `: list o in varlist_attribute' {
            local ++n_geo_attr
            local geo_attr_orig_`n_geo_attr' "`o'"
            local geo_attr_tgt_`n_geo_attr'  "`: word `=`w'+1' of `geo_attr_map''"
            if `n_geo_xw' > 0 local geo_attr_tgt_`n_geo_attr' "`geo_attr_tgt_`n_geo_attr''_attr"
        }
    }

    if "`basic_cleaning'" != "" & (`n_geo_attr' > 0 | `n_geo_xw' > 0) {
        local geo_msg ""
        if `n_geo_attr' > 0 local geo_msg "attribute tables"
        if `n_geo_xw'   > 0 local geo_msg = trim("`geo_msg'" + cond(`n_geo_attr' > 0, " and ", "") + "geo crosswalk (" + trim("`geo_xw_aliases'") + ")")
        di as text "Geography cleaning will use columns from: `geo_msg'"
        if `n_geo_attr' > 0 & `n_geo_xw' > 0 {
            di as text "  Attribute geography columns will get an _attr suffix (e.g. geo_zip_attr)."
        }
    }

    // =========================================================================
    // 4.  PARSE archives() INTO A LIST OF MMYYYY STRINGS
    // =========================================================================

    // Quarter-month mapping
    local q1_mm "03"
    local q2_mm "06"
    local q3_mm "09"
    local q4_mm "12"

    local archive_list ""   // will hold MMYYYY tokens, e.g. 032011

    local arc_input = trim(`"`archives'"')

    // ---- Case D: range of quarter-archives, e.g. "2022Q3-2025Q1" ------------
    if regexm("`arc_input'", "^[0-9]{4}[Qq][1-4]-[0-9]{4}[Qq][1-4]$") {
        local start_part = regexs(0)
        // parse manually
        local dash_pos = strpos("`arc_input'", "-")
        local start_tok = substr("`arc_input'", 1, `dash_pos' - 1)
        local end_tok   = substr("`arc_input'", `dash_pos' + 1, .)

        local sy = substr("`start_tok'", 1, 4)
        local sq = substr("`start_tok'", 6, 1)
        local ey = substr("`end_tok'",   1, 4)
        local eq = substr("`end_tok'",   6, 1)

        local cur_y = `sy'
        local cur_q = `sq'
        local continue 1
        while `continue' {
            local mm = "`q`cur_q'_mm'"
            local archive_list "`archive_list' `mm'`cur_y'"

            // Check if we've reached the end
            if `cur_y' == `ey' & `cur_q' == `eq' {
                local continue 0
            }
            else {
                local cur_q = `cur_q' + 1
                if `cur_q' > 4 {
                    local cur_q 1
                    local cur_y = `cur_y' + 1
                }
            }
        }
    }

    // ---- Case A/B: one or more plain years, e.g. "2021" or "2021 2022" ------
    else if regexm("`arc_input'", "^[0-9]{4}( [0-9]{4})*$") {
        foreach yr of local arc_input {
            foreach q in 1 2 3 4 {
                local mm = "`q`q'_mm'"
                local archive_list "`archive_list' `mm'`yr'"
            }
        }
    }

    // ---- Case B (range of years): e.g. "2021-2023" --------------------------
    else if regexm("`arc_input'", "^[0-9]{4}-[0-9]{4}$") {
        local dash_pos = strpos("`arc_input'", "-")
        local sy = real(substr("`arc_input'", 1, `dash_pos' - 1))
        local ey = real(substr("`arc_input'", `dash_pos' + 1, .))
        if `sy' > `ey' {
            di as error "archives(): start year must be <= end year in a year range."
            exit 198
        }
        forvalues yr = `sy'/`ey' {
            foreach q in 1 2 3 4 {
                local mm = "`q`q'_mm'"
                local archive_list "`archive_list' `mm'`yr'"
            }
        }
    }

    // ---- Case C: specific quarter list, e.g. "2021Q3 2022Q1" ----------------
    else {
        // Try to parse each token as YYYYQq
        foreach tok of local arc_input {
            if regexm("`tok'", "^([0-9]{4})[Qq]([1-4])$") {
                local yr = real(regexs(1))
                local q  = real(regexs(2))
                local mm = "`q`q'_mm'"
                local archive_list "`archive_list' `mm'`yr'"
            }
            else {
                di as error `"archives(): cannot parse token "`tok'". Use YYYY, YYYY-YYYY, YYYYQq, or YYYYQq-YYYYQq formats."'
                exit 198
            }
        }
    }

    local archive_list = trim("`archive_list'")
    if "`archive_list'" == "" {
        di as error "archives(): no valid archives could be parsed from the specification."
        exit 198
    }

    // =========================================================================
    // 5.  VALIDATE OUTPUT DIRECTORY AND FILE NAME
    // =========================================================================

    // Strip trailing slash from dir_output
    local dir_out = trim(`"`dir_output'"')
    if substr("`dir_out'", -1, 1) == "/" | substr("`dir_out'", -1, 1) == "\" {
        local dir_out = substr("`dir_out'", 1, length("`dir_out'") - 1)
    }

    // Strip .dta extension from name_output if supplied
    local fname = trim(`"`name_output'"')
    if lower(substr("`fname'", -4, 4)) == ".dta" {
        local fname = substr("`fname'", 1, length("`fname'") - 4)
    }

    local outpath `"`dir_out'/`fname'.dta"'

    // =========================================================================
    // 5b. NEW (1.3.0): CHECK PARQUET PREREQUISITES (only if parquet requested)
    //     Done before any queries run so a missing tool doesn't surface only
    //     after hours of ODBC pulls.
    // =========================================================================

    if `"`pqsave_options'"' != "" & "`parquet'" == "" {
        di as error "pqsave_options() requires the parquet option."
        exit 198
    }

    if "`parquet'" != "" {

        // (a) Are pqsave.ado, pquse.ado, and pqtools.py on the ado-path?
        local pq_missing ""
        foreach f in pqsave.ado pquse.ado pqtools.py {
            capture findfile `f'
            if _rc local pq_missing "`pq_missing' `f'"
        }
        local pq_missing = trim("`pq_missing'")
        if "`pq_missing'" != "" {
            di as error "The parquet option requires CPL's Parquet tools, but these files were not found on your ado-path: `pq_missing'"
            di as text  "To install them, copy pqsave.ado, pquse.ado, and pqtools.py (all three, into the"
            di as text  "same folder) into your PERSONAL ado directory:"
            di as text  `"    `c(sysdir_personal)'"'
            if `"`cpl_pq_onedev_url'"' != "" {
                di as text `"The files are available on CPL's OneDev server: {browse "`cpl_pq_onedev_url'"}"'
            }
            else {
                di as text "The files are available on CPL's OneDev server (ask a CPL data team member for the repo link)."
            }
            di as text "Or rerun ccp_query without the parquet option to save archive files as .dta."
            exit 601
        }

        // (b) Is Python configured with pyarrow? (pqsave requires Stata 16+
        //     with -python set exec- pointing to a Python that has pyarrow.)
        capture python: import pyarrow
        if _rc {
            di as error "The parquet option requires Python with pyarrow, but Stata could not import pyarrow."
            di as text  "Check {stata python query} and set the right Python with -python set exec-,"
            di as text  "or rerun ccp_query without the parquet option to save archive files as .dta."
            exit 198
        }
    }
	
    // =========================================================================
    // 6.  MAIN QUERY LOOP: iterate over scopes × archives
    // =========================================================================

    // We accumulate all archive-level datasets and append at the end.
    // A temp file list tracks them.
    local tmpfile_list ""
    local arc_count 0

    foreach scope of local scopes {

        // NEW: map the internal scope token to the prefix used in the actual
        // SQL table names (national tables are named e.g. national_032011_attribute).
        // `scope' itself stays "ca"/"ntl" for _archive_scope, messages, and
        // output file names.
        if "`scope'" == "ntl" local tbl_scope "national"
        else                  local tbl_scope "`scope'"

        foreach arc_mm_yr of local archive_list {

            local ++arc_count
            di as text _n "--- Querying archive: `scope'_`arc_mm_yr' (archive `arc_count') ---"

            // ------------------------------------------------------------------
            // 6a.  Build the SQL SELECT for each active source
            // ------------------------------------------------------------------

            // Determine primary (first) source – used as the base table in JOINs
            local primary_src ""
            foreach src in attribute trade inquiry public {
                if `"`varlist_`src''"' != "" & "`primary_src'" == "" {
                    local primary_src "`src'"
                }
            }

            // Build SELECT column list (NEW: `extra_alias_list' appended for flexible join support)
            // CHANGED (1.5.0): built once in section 3b, which also renames
            // extra-table columns whose names clash with columns already pulled
            local select_cols "`select_cols_all'"
            // Remove leading comma
            local select_cols = substr(trim("`select_cols'"), 3, .)

            // Build FROM clause  (primary source)
            local tbl_base "`tbl_scope'_`arc_mm_yr'_`primary_src'"   /* CHANGED: scope -> tbl_scope */
            // NEW (1.2.0): wrap in a derived table if this source has derived fields
            if `"`dsel_`primary_src''"' != "" {
                local tbl_base `"(SELECT * `dsel_`primary_src'' FROM `tbl_base')"'
            }
            local from_clause `"FROM `tbl_base' AS `primary_src'"'

            // Build LEFT JOIN clauses for secondary sources
            local join_clauses ""
            foreach src in attribute trade inquiry public {
                if `"`varlist_`src''"' != "" & "`src'" != "`primary_src'" {
                    local tbl_join "`tbl_scope'_`arc_mm_yr'_`src'"   /* CHANGED: scope -> tbl_scope */
                    // NEW (1.2.0): wrap in a derived table if this source has derived fields
                    if `"`dsel_`src''"' != "" {
                        local tbl_join `"(SELECT * `dsel_`src'' FROM `tbl_join')"'
                    }
                    local on_clause "`sql_on_`primary_src'_`src''"
                    if "`on_clause'" == "" & "`join_key'" != "" {
                        local on_clause "`primary_src'.`join_key' = `src'.`join_key'"
                    }
                    local join_clauses `"`join_clauses' LEFT JOIN `tbl_join' AS `src' ON `on_clause'"'
                }
            }

            // NEW: Build LEFT JOIN clauses for user-specified extra tables
            // (flexible join support). These tables are used verbatim (not
            // archive/scope-prefixed) and joined against primary_src on one
            // or more identically-named fields, ANDed together.
            if `n_extra_tables' > 0 {
                forvalues k = 1/`n_extra_tables' {
                    local tbl_lit "`extra_tblname_`k''"
                    local als     "`extra_alias_`k''"
                    // NEW (1.2.0): wrap in a derived table if this extra table has derived fields
                    if `"`dsel_`als''"' != "" {
                        local tbl_lit `"(SELECT * `dsel_`als'' FROM `tbl_lit')"'
                    }
                    local on_parts ""
                    foreach kk of local extra_keys_`k' {
                        local on_parts "`on_parts' AND `primary_src'.`kk' = `als'.`kk'"
                    }
                    local on_parts = trim("`on_parts'")
                    if substr("`on_parts'", 1, 3) == "AND" {
                        local on_parts = trim(substr("`on_parts'", 4, .))
                    }
                    local join_clauses `"`join_clauses' LEFT JOIN `tbl_lit' AS `als' ON `on_parts'"'
                }
            }

			
            // Build WHERE clause
            local where_parts ""
            foreach src in attribute trade inquiry public {
                if `"`varlist_`src''"' != "" {
                    if `"`sql_where_`src''"' != "" {
                        local where_parts "`where_parts' AND (`sql_where_`src'')"
                    }
                }
            }
            // Add subsample restriction
            if `"`subsample'"' != "" & `"`subsample_where'"' != "" {
                local where_parts "`where_parts' AND (`subsample_where')"
            }
            // NEW (1.4.0): add filter_file() restriction (see section 2c)
            if `has_filter' {
                local where_parts "`where_parts' AND EXISTS (SELECT 1 FROM `filter_tbl' AS `filter_alias' WHERE `filter_cond')"
            }
            // Strip leading AND
            local where_parts = trim("`where_parts'")
            if substr("`where_parts'", 1, 3) == "AND" {
                local where_parts = trim(substr("`where_parts'", 4, .))
            }
            if "`where_parts'" != "" {
                local where_clause "WHERE `where_parts'"
            }
            else {
                local where_clause ""
            }

            // Build IF clause (SQL Server / Sybase style; adapt as needed)
            local if_parts ""
            foreach src in attribute trade inquiry public {
                if `"`varlist_`src''"' != "" & `"`sql_if_`src''"' != "" {
                    local if_parts "`if_parts' `sql_if_`src''"
                }
            }
            local if_clause = trim("`if_parts'")

            // Assemble final SQL
            // CHANGED (1.2.0): compound quotes, so a derived-field expression
            // containing double quotes (e.g. a quoted identifier) can't break
            // the macro assignment
            local sql_query `"SELECT `select_cols' `from_clause' `join_clauses' `where_clause'"'
            if "`if_clause'" != "" {
                local sql_query `"`sql_query' `if_clause'"'
            }

            // ------------------------------------------------------------------
            // 6b.  Execute ODBC query
            // ------------------------------------------------------------------
            //TODO: print query even if it runs!
			capture {
                quietly odbc load, exec(`"`sql_query'"') dsn("`dsn'") clear
            }
            if _rc {
                di as error "ODBC query failed for archive `scope'_`arc_mm_yr'. RC = " _rc
                di as error "SQL attempted:"
                di as text `"`sql_query'"'
                exit _rc
            }
		
            // Add archive identifier variables
            quietly {
                generate str6  _archive_scope = "`scope'"
                generate str6  _archive_mmyyyy = "`arc_mm_yr'"
                // Derive year and quarter for convenience
                local yr4 = substr("`arc_mm_yr'", 3, 4)
                local mm2 = substr("`arc_mm_yr'", 1, 2)
                generate int _archive_year = `yr4'
                local qnum = 0
                if "`mm2'" == "03" local qnum 1
                if "`mm2'" == "06" local qnum 2
                if "`mm2'" == "09" local qnum 3
                if "`mm2'" == "12" local qnum 4
                generate byte _archive_quarter = `qnum'
                label variable _archive_scope   "Archive scope (ca/ntl)"
                label variable _archive_mmyyyy  "Archive month-year (MMYYYY)"
                label variable _archive_year    "Archive year"
                label variable _archive_quarter "Archive quarter (1-4)"
            }

            // ------------------------------------------------------------------
            // 6c.  Optional: basic_cleaning
            // ------------------------------------------------------------------
            //TODO: create a second option to keep the tradeline level data with
			//	cleaning steps (e.g., dedupe tradelines)
			if "`basic_cleaning'" != "" {

				****************************************************************
				* PHASE 1 OF BASIC CLEANING: dropping erroneous record and 
				*		reformatting to save space
				****************************************************************
				
				//Drop "frag" files that have too-low Vantage scores or missing trade-keys
				cap confirm variable vantage_v4_score
				if !_rc {
					cap assert vantage_v4_score < 300
						if _rc qui drop if vantage_v4_score < 300
				}
					
				//Drop deceased people
				cap confirm variable deceased_in
				if !_rc {
					cap assert deceased_in == "N"
						if _rc qui drop if deceased_in == "Y"
						drop deceased_in
				}
					
				//Destring and reformat birth date
				cap confirm variable birth_dt_cust
				if !_rc {
					
					//missing birthdates are signs of frag files (2010Q3 onwards)
					qui drop if mi(birth_dt_cust) & (`yr4' > 2010 | (`yr4' == 2010 & real("`mm2'") >= 9))
						
					//impute January birth for people w/o month (month = 00)
					qui replace birth_dt_cust = subinstr(birth_dt_cust,"00","01",1) if ///
							strpos(birth_dt_cust,"00") == 1
								
					//reformat birth date (should never be missing if after Q2 2010)
					qui gen long person_birth_date = date(birth_dt_cust,"MY")
						drop birth_dt_cust
						format person_birth_date %td
				}
				
				
				//Destring any other dates (variables ending in date/dt)
				foreach suffix in date dt {
					
					//check if any variables of this type exist
					qui cap unab vars : *`suffix' 
					if _rc continue 
					
					//if they do, destring and reformate to save space
					foreach var of varlist *`suffix' {
							
						//break if there isn't a date variable of the given format, or if the variable is birth_date
						cap confirm variable `var'
						if (_rc != 0 | regexm("`var'","birth_date")) continue 
							
						//otherwise, rename the variable "`var'_date" and reformat as variable with _int suffix 
						local newname = trim(subinstr("`var'","_`suffix'","_date",1))
						qui gen int `newname'_int = date(`var',"YMD")
						format `newname'_int %td
						drop `var'
						ren `newname'_int `newname'
					}
				}
								
				//Drop duplicate tradelines, if possible/necessary 
				
					*only necessary if the data include tradelines
					if `incl_tradeline' == 1 {
						
						//variables we need to identify unique tradelines
						local tradeids ctk file_date acct_open_dt acct_type_cd enhanced_acct_type_cd terms_trans terms_freq_cd ///
							lender_key amount_1 amount_2 acct_balance_am acct_past_due_am acct_payment_am ///
							actual_pymt_am acct_credit_limit_am acct_high_balance_am acct_dollar_am ///
							acct_charge_off_am acct_status_cd enhanced_acct_status_cd ///
							is_open payment_grid ecoa_cd
							
						//confirm that these tradeline identifiers are present
						//TODO: add logic to deduplicate on whatever variables
						//	are present and print out the variables that are 
						//	missing
						local has_tradeids = 1
						local mi_tradeids_list ""
						foreach var of local tradeids {
							cap confirm variable `var'
							if _rc {
								local mi_tradeids_list "`mi_tradeids_list' `var'"
								local has_tradeids = 0
							}
						}
						
						//drop duplicate tradelines if we have the variables to do so
						if `has_tradeids' == 1 {
							
							*drop duplicate tradelines
							bys `tradeids': gen byte non_dupe_trade = 1 if _n == 1
								qui drop if non_dupe_trade != 1
								drop non_dupe_trade
							
						}
						
						//otherwise, inform the user that we can't de-duplicate tradelines
						if `has_tradeids' == 0 & mi(`no_tradeids') {
							di ""
							di as text "NOTE: we tried to remove duplicate tradelines, but not enough information was provided. Please include the variables: "
							di as text "		`mi_tradeids_list'"
							di as text "in your driver file if you want to de-duplicate tradelines."
							local no_tradeids = 1
						}
					}
					
				****************************************************************
				* PHASE 2 OF BASIC CLEANING: renaming attribute data to clearly 
				*		  indicate that the data are at the person-quarter level; 
				*		  rename geographic variables to signal they are geographies
				****************************************************************
				
				* Typical demographic variables
				local demographics marriage_in gender_code hshld_child_ct
				foreach var of local demographics {
					cap confirm variable `var'
					if !_rc ren `var' person_`var'
				}
				
				* Geographies 
				* CHANGED (1.5.0): source-aware. Which columns came from the
				*	attribute tables and which from a geo crosswalk, and the
				*	name each gets, are worked out in section 3b.
				
					//(a) geography from the attribute tables: rename for
					//	consistency (geo_zip, geo_stabbr, ..., with an _attr
					//	suffix if a geo crosswalk is also pulled), then set
					//	codes that are not numeric to missing.
					//	FIX (1.5.0): the old check compared each variable's
					//	first value, not its name, to "geo_stabbr".
					forvalues j = 1/`n_geo_attr' {
						cap confirm variable `geo_attr_orig_`j'', exact
						if _rc continue
						ren `geo_attr_orig_`j'' `geo_attr_tgt_`j''
						if "`geo_attr_orig_`j''" == "state_cd" continue
						cap confirm string variable `geo_attr_tgt_`j''
						if !_rc qui replace `geo_attr_tgt_`j'' = "" if mi(real(`geo_attr_tgt_`j''))
					}
					
					//(b) geography from a geo crosswalk (e.g.
					//	addr_geo_xwalk_v2): rename only. Its codes are
					//	validated when the crosswalk is built, and school
					//	and district IDs must keep their leading zeros.
					forvalues j = 1/`n_geo_xw' {
						if "`geo_xw_recv_`j''" == "`geo_xw_tgt_`j''" continue
						cap confirm variable `geo_xw_recv_`j'', exact
						if _rc continue
						cap confirm variable `geo_xw_tgt_`j'', exact
						if !_rc {
							if `arc_count' == 1 di as text "NOTE: kept `geo_xw_recv_`j'' as is; `geo_xw_tgt_`j'' already exists."
							continue
						}
						ren `geo_xw_recv_`j'' `geo_xw_tgt_`j''
					}
					
				*Person-level variables 
					
					//credit score
					cap confirm variable vantage_v4_score
						if !_rc ren vantage_v4_score person_credit_score
					
					//race/ethnicity
					qui cap unab vars: pr_* 
					if !_rc {
						ren pr_* person_race_*
					}
						
				****************************************************************
				* PHASE 3 OF BASIC CLEANING: beginning to aggregate tradeline data
				****************************************************************
				
				* This step is only relevant if we're using tradeline data
				if `incl_tradeline' == 1 { 
				
					* Payment history and account age
				
						//flag tradelines that are open; an account that is open andis not collection/debt buyer --> open credit 
						cap confirm variable acct_cond_cd loan_type
						if !_rc {
							qui gen byte credit_num_open_unw = cond(acct_cond_cd == "A1" & !inlist(loan_type,"Collection","Debt buyer"),1,0)
						} 
						
						//length of payment history (+ on open accounts)
						cap confirm variable payment_grid_84
						if !_rc {
							qui gen byte credit_payment_length = length(payment_grid_84)
						}
						
						//length of open tradeline payment histories
						// NOTE: "credit_num_open_unw" is a binary!
						cap confirm variable credit_num_open_unw credit_payment_length
						if !_rc {
							qui gen byte credit_payment_length_open = credit_payment_length * credit_num_open_unw
						}
						
						//age of accounts (+ open accounts) --> credit age in years
						// NOTE: "credit_num_open_unw" is a binary!
						cap confirm variable file_date acct_open_date
						if !_rc {
							qui gen credit_age = (file_date - acct_open_date)/365.25
						}
						cap confirm variable file_date acct_open_date credit_num_open_unw
						if !_rc {
							qui gen credit_age_open = credit_age * credit_num_open_unw
						}
						
					*Initialize measure of total # of tradelines 
					cap confirm variable trade_key
					if !_rc qui gen int credit_num_alltrades = !mi(trade_key)
					
					*drop variables that we no longer need 
					foreach var in acct_open_dt terms_trans terms_freq_cd ///
						lender_key amount_1 amount_2 acct_past_due_am acct_payment_am ///
						actual_pymt_am acct_high_balance_am acct_charge_off_am payment_grid ///
						acct_status_cd enhanced_acct_status_cd {
						
							cap confirm `var'
							if !_rc drop `var'
					}
			
				
				//Prepare person-quarter tradeline variables 
				
					*Check if this is necessary -- need to have loan types to 
					*	make sense of the tradeline data
					cap confirm variable loan_type
					if _rc & mi("`no_loantype'") {
						di ""
						di as text "NOTE: we tried to aggregate the tradeline data to the person-quarter level. However, we noticed that loan_type is missing."
						di as text "	Without loan_type, we can't determine what kind of account a tradeline refers to and can't aggregate across tradelines."
						di ""
						local no_loantype = 1
					} 
					
					if !_rc {
						
						//assign weights to tradelines (i.e., adjust for joint accounts)
						cap confirm variable ecoa_cd
						if _rc {
							
							*tell the user that we can't assign weights to joint accounts
							if mi("`no_ecoacd'") {
								di ""
								di as text "NOTE: we tried to weight the tradeline data to adjust for joint accounts. However, we noticed that ecoa_cd is missing."
								di as text "	Without ecoa_cd, we can't determine which accounts are joint."
								di ""
								local no_ecoacd = 1
							}
							
							*assign a dummy "weight" that's always 1
							qui gen byte weight = 1
							
								//data are unweighted 
								local wgt = "unw"
						}
						
						if !_rc {
							
							*default to zero weights
							qui gen weight = 0
							
								//note that we can create weighted data
								local wgt = "wgt"
								
								//full weight to primary accountholders
								qui replace weight = 1 if inlist(real(ecoa_cd),1,7)
									
								//half-weight to debt from joint accounts
								qui replace weight = 0.5 if inlist(real(ecoa_cd),2,4)
						
								//TODO: check that 3 is NOT authorized user/check what 3 is
								//all others [ecoa_cds 3, 5, and 6, which are guarantor-type accounts, as well as closed accounts] have zero-weight by default	
							
							*drop ecoa_cd
							drop ecoa_cd
							
							*create a number of weighted credit lines
							qui gen byte credit_num_open_wgt = credit_num_open_unw * weight
						}
					
						
						*Define loan types that we care about
						local loantypes "creditcard homeloan homeequity studentloan autoloan otherloan collection" 
						
							//credit cards
							local creditcard = `""Retail credit card", "Credit card""'

							//home loans
							local homeloan = `""Mortgage""'
							
							//home equity 
							local homeequity = `""Home equity""'
							
							//student loans
							local studentloan = `""Student loan""'
							
							//auto loans
							local autoloan = `""Auto lease","Auto loan""'
							
							//other
							local otherloan = `""Miscellany","Other secured","Unsecured loan","Utilities""'
							
							//collections
							local collection = `""Debt buyer","Collection""' 
			
		
						*Loan balances, limits, and # of lines by type, for open accounts and primary holders
						foreach ty of local loantypes {

							if inlist("`ty'", "collection") {
								local open ""
							}
							
							// the "open" field is not relevant for collection items
							if !inlist("`ty'","collection") {
								local open "& is_open" 
							}
						
							cap confirm variable is_open 
							if !_rc {
								
								//credit balances (weighted only)
								cap confirm variable acct_balance_am 
								if !_rc qui gen credit_tot_`ty'_amt_`wgt' = acct_balance_am * weight if inlist(loan_type,``ty'') `open' 
								
								// when collapse these, will be sum (total # ) of credit lines/loans
								qui gen int credit_num_`ty'_open_unw = 1 if inlist(loan_type,``ty'') `open' 
								if "`wgt'" == "wgt" {
									qui gen int credit_num_`ty'_open_`wgt' = weight if inlist(loan_type,``ty'') `open' 
								}
								
								//percent of principal paid 
								cap confirm variable acct_balance_am acct_dollar_am
								if !_rc qui gen credit_`ty'_pctp_`wgt' = (acct_balance_am * weight) / acct_dollar_am if inlist(loan_type,``ty'') `open' 
								
								//next two for credit cards and home equity ONLY
								if inlist("`ty'","creditcard","homeequity") {
									
									cap confirm variable acct_credit_limit_am

									//limit  (weighted)
									if !_rc qui gen credit_tot_`ty'_limit_`wgt' = acct_credit_limit_am * weight if inlist(loan_type,``ty'')  `open' 
									
									//credit utilization (credit cards and home equity only)
									// this is avg. --> will also create total credit utilization as [(sum) of balance for given type]/[ (sum) of acct_credit_limit]
									if !_rc qui gen credit_`ty'_util_`wgt' = (acct_balance_am * weight) /acct_credit_limit_am if inlist(loan_type,``ty'') `open'
								}
							}
						}
					}
					
					//Collapse the data to create a person-quarter dataset 
					
						* Collapse data to person-quarter level
						* mean util (perc_princ) is the avg. utilization (percentage of principal paid) across accounts
						* sum: total balance across loan type, total credit limit across loan type, total number of accounts/open accounts/credit lines 
						* max: demographics, credit score (only 1, so same as mean), take oldest credit age /length of payment history 
						* geo is consistent within ctk by quarter, so no duplicates 
					
							//construct arguments for the collapse step 
							
								*take the mean of utilization and pct principal paid
								local args_mean_list = "(mean)"
								foreach ty in util pctp {
									qui cap unab vars : *_`ty'_
									if !_rc {
										local args_mean_list = "`args_mean_list' *_`ty'_*"
									}
								}
								
								*take the sum of credit_tot* and credit_num* variables
								local args_sum_list = "(sum)"
								foreach ty in credit_tot credit_num {
									qui cap unab vars : `ty'*
									if !_rc {
										local args_sum_list = "`args_sum_list' `ty'*"
									}
								}
								
								*take the max of person, principal, payment, and credit age variables
								local args_max_list = "(max)"
								foreach ty in credit_payment credit_ag {
									qui cap unab vars : `ty'*
									if !_rc {
										local args_max_list = "`args_max_list' `ty'*"
									}
								}
						
								*see what "by" variables there are 
								local args_by_list = ""
								foreach ty in person geo {
									qui cap unab vars : `ty'_*
									if !_rc {
										local args_by_list = "`args_by_list' `ty'_*"
									}
								}
								
								*reset any lists that are empty
								foreach ty in mean sum max {
									if "`args_`ty'_list'" == "(`ty')" local args_`ty'_list = ""
								}
								
								*edge case: if there's nothing to collapse, then 
								*	skip that step and display a message to the 
								*	user
								local collapse_cmd = "`args_mean_list' `args_sum_list' `args_max_list'"
								
							//collapse the data if there's actually data to collapse
							if trim("`collapse_cmd'") != "" {
								gcollapse `args_mean_list' `args_sum_list' `args_max_list', ///
									  by(ctk _archive_* `args_by_list')
									  
								gisid ctk _archive_mmyyyy
							}
							
							else {
								if mi("`no_collapse'") {
									di ""
									di as text "NOTE: we tried to collapse the tradeline data to the person-quarter level. However, we could not construct typical quarterly tradeline variables"
									di as text "	given the specified data. Try including common fields like loan_type and acct_balance_am, or else collapse the data using a bespoke approach."
									di as text "	The data are NOT unique by person-quarter."
									di ""
								}
								
								local no_collapse = 1
							}
				}

				
				****************************************************************
				* END OF BASIC CLEANING: wrap up this quarterly file 
				****************************************************************
                di as text "  [basic cleaning steps applied to `scope'_`arc_mm_yr']"
            }
		
            // Save quarterly file
			local fi_name = "build_ccp_`scope'_`mm2'`yr4'_`subsample'pct"
			qui compress
            quietly save "`dir_out'/`fi_name'", replace
            local tmpfile_list "`tmpfile_list' `fi_name'"

            // NEW (1.3.0): also write this archive as Parquet. The .dta just
            // saved is unchanged in memory, so pqsave reads it directly
            // (no extra copy). The .dta is kept until the final append below
            // succeeds, then erased (section 10). If the Parquet write fails,
            // the run continues and that archive's .dta is kept instead.
            if "`parquet'" != "" {
                capture noisily pqsave, name("`fi_name'") directory("`dir_out'") replace `pqsave_options'
                if _rc {
                    di as error "  WARNING: Parquet save failed for `fi_name' (rc = " _rc "); keeping `fi_name'.dta instead."
                    local pq_failed_list "`pq_failed_list' `fi_name'"
                }
            }
		}
	}

    // =========================================================================
    // 7.  APPEND ALL ARCHIVE DATASETS
    // =========================================================================

    if `arc_count' == 0 {
        di as error "No archives were queried successfully."
        exit 198
    }

    // Load first archive, then append the rest
    local files_arr : list sizeof tmpfile_list
    local first_file : word 1 of `tmpfile_list'
    quietly use "`dir_out'/`first_file'", clear

    if `arc_count' > 1 {
        forvalues i = 2/`arc_count' {
            local next_file : word `i' of `tmpfile_list'
            quietly append using "`dir_out'/`next_file'"
        }
    }

    //di as text _n "Total rows after appending all archives: " _N
	di as text "Total archives queried and appended: " `arc_count'

    // =========================================================================
    // 8.  OPTIONAL: APPEND TO EXISTING .DTA FILE
    // =========================================================================

    if `"`append_existing'"' != "" {
        local ae_path = trim(`"`append_existing'"')
        capture confirm file `"`ae_path'"'
        if _rc {
            di as error `"append_existing(): file not found: `ae_path'"'
            exit 601
        }
        quietly append using `"`ae_path'"'
        di as text "Appended to existing file: `ae_path'"
        di as text "Total rows after appending existing file: " _N
    }

    // =========================================================================
    // 9.  SAVE OUTPUT
    // =========================================================================
	qui compress
    quietly save `"`outpath'"', replace
    di as text _n "Output saved to: `outpath'"

    // =========================================================================
    // 10. NEW (1.3.0): PARQUET CLEANUP
    //     The quarterly .dta files were only needed for the append above;
    //     remove those that were successfully written as Parquet. Never
    //     erases the combined output file.
    // =========================================================================

    if "`parquet'" != "" {
        local n_pq_ok 0
        foreach f of local tmpfile_list {
            local f_failed : list f in pq_failed_list
            if !`f_failed' & `"`dir_out'/`f'.dta"' != `"`outpath'"' {
                capture erase "`dir_out'/`f'.dta"
                local ++n_pq_ok
            }
        }
        di as text "Quarterly archive files saved as Parquet in `dir_out': `n_pq_ok'"
        local pq_failed_list = trim("`pq_failed_list'")
        if "`pq_failed_list'" != "" {
            di as error "Parquet save failed for (kept as .dta): `pq_failed_list'"
        }
        di as text "Read an archive back into Stata with: pquse <file>.parquet, directory(`dir_out') clear"
    }

end
