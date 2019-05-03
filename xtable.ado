/*---------------------------------------------------------------------		
/// XTABLE - export table output to excel

/// Weverthon Machado

v0.3.1 - 2019-05-03
---------------------------------------------------------------------*/
capture program drop xtable
program define xtable
version 13,1

/* Get only -table- options that are needed here. The final asterisk 
captures everything else */
syntax varlist(max=3) [if] [in] [fw aw pw iw] [, /* 
		*/ BY(varlist) COLumn REPLACE ROW SColumn  *]


/*********************************************************************
# Parse arguments
**********************************************************************/
/* Tokenize */
tokenize
local tableargs "`0'"

tokenize `varlist'
local rowvar = "`1'"
local colvar = "`2'"
local scolvar = "`3'"
local srowvar =  "`by'"


/* Run -table- */
preserve
table `tableargs' replace


/* Get numbers and levels of vars and stats */
qui levelsof `rowvar', local(row_levels) missing
local nrow: list sizeof row_levels

if !missing("`colvar'") {
	qui levelsof `colvar', local(col_levels) missing
	local ncol: list sizeof col_levels
}
else {
	local col_levels 1
	local ncol = 1
}


if !missing("`scolvar'") {
	qui levelsof `scolvar', local(scol_levels) missing
	local nscol: list sizeof scol_levels
}
else {
	local scol_levels 1
	local nscol = 1
}


if !missing("`srowvar'") {
	qui levelsof `srowvar', local(srow_levels) missing
	local nsrow: list sizeof srow_levels
}
else {
	local srow_levels 1
	local nsrow = 1
}


unab stat_list: table*
local nstats: word count `stat_list'


if !(missing("`colvar'") & missing("`srowvar'")) {
	cap drop _fillin
	fillin `rowvar' `colvar' `scolvar' `srowvar'
	drop _fillin
}

/*********************************************************************
# Build matrix
**********************************************************************/
matrix xt_results = J(1,`ncol'*`nscol', .)

foreach srow in `srow_levels' {

	if `nsrow' > 1 {
		tempfile stats_data
		qui save `stats_data'
		qui keep if `srowvar' == `srow'
	}

	local psrow: list posof "`srow'" in srow_levels
	mat def xt_scol_`psrow' = J(`nrow'*`nstats', 1, .)

	foreach scol in `scol_levels' {

		if `nscol' > 1 {
		tempfile stats_data_`srow'
		qui save `stats_data_`srow''
		qui keep if `scolvar' == `scol'
		}

		local pscol: list posof "`scol'" in scol_levels
		mat def xt_`pscol' = J(`nrow'*`nstats', `ncol', .)

		sort `rowvar' `colvar' `scolvar' `srowvar'
		forvalues row = 1/`nrow' {
			forvalues col = 1/`ncol' {
				forvalues stat = 1/`nstats' {

						mat xt_`pscol'[((`row'-1)*`nstats')+`stat', `col'] = table`stat'[((`row'-1)*`ncol')+`col']

				}
			}
		}


		mat xt_scol_`psrow' = xt_scol_`psrow',  xt_`pscol'
		mat drop xt_`pscol'

		if `nscol' > 1 {
			qui use `stats_data_`srow'', clear
		}
	}

	mat xt_scol_`psrow' = xt_scol_`psrow'[1..., 2...]

	if `nsrow' > 1 {
		mat srow_header = J(1,`ncol'*`nscol', .)
		mat xt_results = xt_results \ srow_header \ xt_scol_`psrow'
	}
	else {
		mat xt_results = xt_results \ xt_scol_`psrow'
	}
	mat drop xt_scol_`psrow'

	if `nsrow' > 1 {
		qui use `stats_data', clear
	}
}


mat xtable = xt_results[2..., 1...]
mat drop xt_results


/*********************************************************************
# Labels
**********************************************************************/

/* Rows */
foreach srow in `srow_levels' {

	local psrow: list posof "`srow'" in srow_levels

	foreach row in `row_levels' {
		local row_label : label (`rowvar') `row'
		#delimit ;
		local mat_rownames_`psrow' = `"`mat_rownames_`psrow''"' + `" ""' + 
									subinstr(substr("`row_label'", 1, 30), ".", " ", .) + `"""' +
								    (`""-""')*(`nstats'-1)
		;
		#delimit cr
		
	}

	if `nsrow' > 1 {
		local srow_label : label (`srowvar') `srow'
		local mat_rownames_`psrow' =  `" ""' + subinstr(substr("`srow_label'", 1, 30), ".", " ", .) + `"""' + `"`mat_rownames_`psrow''"'
	}

	local mat_rownames = `"`mat_rownames'"' + `"`mat_rownames_`psrow''"' 
}

mat rownames xtable = `mat_rownames'

/* Columns */
if !missing("`colvar'") {

	foreach scol in `scol_levels' {

		local pscol: list posof "`scol'" in scol_levels

		foreach col in `col_levels' {
			local col_label : label (`colvar') `col'
			#delimit ;
			local mat_colnames_`pscol' = `"`mat_colnames_`pscol''"' + `" ""' + 
										subinstr(substr("`col_label'", 1, 30), ".", " ", .) + `"""'
			;
			#delimit cr
			
		}

		local mat_colnames = `"`mat_colnames'"' + `"`mat_colnames_`pscol''"' 
	}

mat colnames xtable = `mat_colnames'

}


/*********************************************************************
# Export
**********************************************************************/
local rowvar_label: var label `rowvar'

#delimit ;
qui putexcel A2 = matrix(xtable, names) A2 = ("`rowvar_label'")
			 using xtable.xlsx, replace
;
#delimit cr

if !missing("`colvar'") {
	local colvar_label: var label `colvar'
	qui putexcel B1 = ("`colvar_label'") using xtable.xlsx, modify
}

di as smcl "Output written to {browse  "`"xtable.xlsx}"'" 
end
