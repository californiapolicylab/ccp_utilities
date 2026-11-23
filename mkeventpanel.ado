*! mkeventpanel version 3.1.0  23nov2026
*! Creates an event-time panel from individual-by-quarter data.
*! Uses Stata frames (v16+); no preserve/restore; minimal peak memory.
*
* ============================================================
* SYNTAX
* ============================================================
*
*   mkeventpanel, id(varname) time(varname) event(varname)
*                 pre(#) post(#)
*                 [keepvars(varlist) basevars(varlist)]
*
* REQUIRED OPTIONS
*   id(varname)      Person identifier (numeric or string).
*   time(varname)    Quarter variable: string ("2014Q4"/"2014q4"),
*                    %tq date, or integer (value labels preserved).
*   event(varname)   0/1 event indicator (missing treated as 0).
*                    Each 1 produces its own person-event panel.
*   pre(#)           Pre-event quarters to include (>= 0).
*   post(#)          Post-event quarters to include (>= 0).
*
* OPTIONAL OPTIONS
*   keepvars(varlist)  Variables to carry into the panel.  id() and
*                      time() are always kept.  If omitted, all
*                      variables are kept.
*   basevars(varlist)  Creates time-invariant b_<var> equal to <var>
*                      in the quarter before the event (event time -1).
*                      Names are truncated to 32 characters.  Any
*                      basevars missing from keepvars() are added to
*                      it automatically.  Requires pre() >= 1.
*
* OUTPUT VARIABLES
*   _pe_id        Integer ID, unique per person x event.
*   _event_num    Within-person event sequence (1 = earliest).
*   _event_time   Quarters relative to event (0 = event quarter).
*   b_<var>       Baseline value for each basevars() variable.
*
* The result is unique at (_pe_id, _event_time), balanced, and
* sorted by (_pe_id, _event_time).  Quarters in a window with no
* source observation appear with missing covariates.
*
* WARNING: The data in memory are replaced by the event panel, and
* the original data are cleared early in the run to save memory.
* Keep a copy of the source data on disk.
*
* REQUIRES Stata 16 or later (frames).
* ============================================================

program define mkeventpanel
    version 16.0

    syntax ,                    ///
        id(varname)             ///
        time(varname)           ///
        event(varname)          ///
        pre(integer)            ///
        post(integer)           ///
        [ keepvars(varlist)     ///
          basevars(varlist)     ]

    /* ----------------------------------------------------------
       0.  Validate inputs (before anything is modified)
    ---------------------------------------------------------- */

    if `pre' < 0 | `post' < 0 {
        di as error "pre() and post() must be >= 0."
        exit 198
    }
    if `pre' == 0 & `post' == 0 {
        di as error "pre() and post() cannot both be 0."
        exit 198
    }
    if "`basevars'" != "" & `pre' == 0 {
        di as error "basevars() requires pre() >= 1."
        exit 198
    }
    capture assert inlist(`event', 0, 1, .)
    if _rc {
        di as error "event(`event') must be a 0/1 binary variable."
        exit 198
    }

    local timetyp : type `time'
    local timefmt : format `time'
    local timelbl : value label `time'
    local isstr = substr("`timetyp'", 1, 3) == "str"

    if `isstr' {
        tempvar _yr _qt
        quietly gen int `_yr' = real(substr(strtrim(`time'), 1, 4))
        quietly gen int `_qt' = real(substr(strtrim(`time'), -1, 1))
        capture assert !missing(`_yr') & inrange(`_qt', 1, 4)
        if _rc {
            drop `_yr' `_qt'
            di as error "Cannot parse time() as YYYYQq (e.g. 2014Q4)."
            exit 198
        }
        drop `_yr' `_qt'
    }

    quietly count if `event' == 1
    local nevents = r(N)
    if `nevents' == 0 {
        di as error "No observations with event() == 1."
        exit 2000
    }

    // Fold basevars into keepvars so the user need not list them twice
    if "`keepvars'" != "" & "`basevars'" != "" {
        local keepvars : list keepvars | basevars
    }
    local keepvars : list keepvars - id
    local keepvars : list keepvars - time
    local keepvars : list keepvars - event

    /* ----------------------------------------------------------
       1.  Baseline output names: "b_" + name, capped at 32 chars
    ---------------------------------------------------------- */

    local basenames ""
    if "`basevars'" != "" {
        local _trunc_any 0
        foreach _bv of local basevars {
            local _raw = "b_" + "`_bv'"
            local _out = substr("`_raw'", 1, 32)
            if "`_out'" != "`_raw'" {
                if !`_trunc_any' {
                    di as text "(mkeventpanel) b_ name(s) truncated to 32 characters:"
                    local _trunc_any 1
                }
                di as text "    `_raw'  ->  `_out'"
            }
            local basenames = strtrim("`basenames'" + " " + "`_out'")
        }

        // Report any output names that collide after truncation
        local _unames : list uniq basenames
        if `: word count `_unames'' != `: word count `basenames'' {
            di as error "basevars(): output names collide after truncation:"
            local i 1
            foreach _bn of local basenames {
                local j 1
                foreach _bn2 of local basenames {
                    if `j' > `i' & "`_bn'" == "`_bn2'" {
                        local _v1 : word `i' of `basevars'
                        local _v2 : word `j' of `basevars'
                        di as error "    `_bn'  <-  `_v1' and `_v2'"
                    }
                    local ++j
                }
                local ++i
            }
            exit 198
        }

        // Make sure new names will not clash with variables that will
        // be in the output
        foreach _bn of local basenames {
            if "`keepvars'" != "" {
                if `: list _bn in keepvars' {
                    di as error "basevars(): output name `_bn' already exists in keepvars()."
                    exit 110
                }
            }
            else {
                capture confirm variable `_bn', exact
                if !_rc {
                    di as error "basevars(): variable `_bn' already exists in the data."
                    exit 110
                }
            }
        }
    }

    /* ----------------------------------------------------------
       2.  Build _mep_src with ONLY the needed columns, then clear
           the original data to free memory.
    ---------------------------------------------------------- */

    local curframe = c(frame)
    foreach _f in _mep_src _mep_spine _mep_base {
        capture frame drop `_f'
    }

    // Preserve the time variable's value-label definition (ecode case)
    if "`timelbl'" != "" {
        tempfile _lblfile
        quietly label save `timelbl' using `_lblfile'
    }

    if "`keepvars'" != "" {
        frame put `id' `time' `event' `keepvars', into(_mep_src)
    }
    else {
        frame put _all, into(_mep_src)
    }

    // Original data are no longer needed; free them now
    quietly drop _all

    frame _mep_src {

        // Compact integer quarter index _tq
        if `isstr' {
            gen long _tq = (real(substr(strtrim(`time'), 1, 4)) - 1960) * 4 ///
                         + real(substr(strtrim(`time'), -1, 1)) - 1
        }
        else {
            gen long _tq = `time'
        }
        quietly compress _tq

        // String time no longer needed in the source (rebuilt at end);
        // numeric time duplicates _tq.  Drop it either way.
        drop `time'

        // Event roster: one row per person-event
        frame put `id' _tq if `event' == 1, into(_mep_spine)
        drop `event'
    }

    /* ----------------------------------------------------------
       3.  Event roster -> IDs -> (optional) baseline lookup
    ---------------------------------------------------------- */

    frame _mep_spine {
        rename _tq _ev_root
        sort `id' _ev_root
        by `id': gen int _event_num = _n
        gen long _pe_id = _n
        quietly compress _event_num _pe_id
    }

    if "`basevars'" != "" {

        // One row per person-event; key = quarter before the event
        frame _mep_spine: frame put `id' _pe_id _ev_root, into(_mep_base)

        frame _mep_base {
            gen long _tq = _ev_root - 1
            drop _ev_root

            // frlink uses hashing; no sort required
            frlink m:1 `id' _tq, frame(_mep_src)

            local i 1
            foreach _bv of local basevars {
                local _bn : word `i' of `basenames'
                frget `_bn' = `_bv', from(_mep_src)
                local ++i
            }
            drop _mep_src _tq `id'
        }
    }

    /* ----------------------------------------------------------
       4.  Expand to the event window
    ---------------------------------------------------------- */

    frame _mep_spine {
        local winlen = `pre' + `post' + 1
        quietly expand `winlen'
        sort _pe_id
        by _pe_id: gen int _event_time = (_n - 1) - `pre'
        gen long _tq = _ev_root + _event_time
        drop _ev_root
        quietly compress _event_time _tq
    }

    /* ----------------------------------------------------------
       5.  Attach baselines (by _pe_id), then covariates (by id, _tq)
    ---------------------------------------------------------- */

    if "`basevars'" != "" {
        frame _mep_spine {
            frlink m:1 _pe_id, frame(_mep_base)
            frget `basenames', from(_mep_base)
            drop _mep_base
        }
        frame drop _mep_base
    }

    frame _mep_src {
        local _srcvars : varlist _all
    }
    local _pull : list _srcvars - id
    local _tqname _tq
    local _pull : list _pull - _tqname

    frame _mep_spine {
        frlink m:1 `id' _tq, frame(_mep_src)
        if "`_pull'" != "" {
            frget `_pull', from(_mep_src)
        }
        drop _mep_src
    }

    frame drop _mep_src

    /* ----------------------------------------------------------
       6.  Rebuild time variable, label, order, sort
    ---------------------------------------------------------- */

    frame _mep_spine {

        if `isstr' {
            gen str7 `time' = string(floor(_tq / 4) + 1960) + "Q" ///
                            + string(mod(_tq, 4) + 1)
            drop _tq
        }
        else {
            rename _tq `time'
            capture recast `timetyp' `time'
            format `time' `timefmt'
            if "`timelbl'" != "" {
                quietly do `_lblfile'
                label values `time' `timelbl'
            }
        }

        label variable _pe_id      "Person-event ID"
        label variable _event_num  "Within-person event sequence (1 = first)"
        label variable _event_time "Quarters since event (0 = event quarter)"

        if "`basevars'" != "" {
            local i 1
            foreach _bv of local basevars {
                local _bn : word `i' of `basenames'
                label variable `_bn' "`_bv' in quarter before event"
                local ++i
            }
        }

        order `id' _pe_id _event_num _event_time `time' `basenames'
        sort _pe_id _event_time

        quietly count
        local nrows = r(N)
        quietly count if _event_num == 1 & _event_time == 0
        local npersons = r(N)
    }

    /* ----------------------------------------------------------
       7.  Swap the result into the user's frame (no copy)
    ---------------------------------------------------------- */

    frame change _mep_spine
    frame drop `curframe'
    frame rename _mep_spine `curframe'

    /* ----------------------------------------------------------
       8.  Report
    ---------------------------------------------------------- */

    di as text ""
    di as text "  {hline 52}"
    di as text "  {bf:mkeventpanel} complete"
    di as text "  {hline 52}"
    di as text "  Persons with an event    : " as result %12.0fc `npersons'
    di as text "  Person-event panels      : " as result %12.0fc `nevents'
    di as text "  Window                   : " as result "-`pre'" ///
               as text " to " as result "+`post'" as text " quarters"
    di as text "  Observations             : " as result %12.0fc `nrows'
    if "`basevars'" != "" {
        di as text "  Baseline variables       : " as result "`basenames'"
    }
    di as text "  {hline 52}"

end
