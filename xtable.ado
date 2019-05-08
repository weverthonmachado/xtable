/*---------------------------------------------------------------------		
/// XTABLE - export table output to excel

/// Weverthon Machado

v0.5 - 2019-05-08
---------------------------------------------------------------------*/
program define xtable, rclass
version 13.1

/* Get only -table- options that are needed here. The final asterisk 
captures everything else */
syntax varlist(max=3) [if] [in] [fw aw pw iw] [, /* 
		*/ BY(varlist) REPLACE *]


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

tokenize `by'
local srow1var = "`1'"
local srow2var = "`2'"
local srow3var = "`3'"
local srow4var = "`4'"
local nby: list sizeof by	



/* Run -table- */
preserve
if !missing("`replace'") {
	table `tableargs'
}
else {
	if !missing("`options'") | (missing("`options'") & regexm("`tableargs'", ",")) {
		table `tableargs' replace
	}
	else {
		table `tableargs', replace
	}
}

/* Get numbers and levels of vars and stats */
qui levelsof `rowvar', local(row_levels) missing
local nrow: list sizeof row_levels

unab stat_list: table*
local nstats: word count `stat_list'

if !missing("`colvar'") {
	qui levelsof `colvar', local(col_levels) missing
	local ncol: list sizeof col_levels
}
else {
	local ncol = `nstats'
}


if !missing("`scolvar'") {
	qui levelsof `scolvar', local(scol_levels) missing
	local nscol: list sizeof scol_levels
}
else {
	local scol_levels 1
	local nscol = 1
}

forvalues n = 1/4 {
	if !missing("`srow`n'var'") {
		qui levelsof `srow`n'var', local(srow`n'_levels) missing
		local nsrow`n': list sizeof srow`n'_levels
	}
	else {
		local srow`n'_levels 1
		local nsrow`n' = 1
	}
}


if !(missing("`colvar'") & missing("`srow2var'")) {
	cap drop _fillin
	fillin `rowvar' `colvar' `scolvar' `srow1var' `srow2var' `srow3var' `srow4var'
	drop _fillin
}


/*********************************************************************
# Build matrix
**********************************************************************/

matrix xt_results = J(1,`ncol'*`nscol', .)

/*-----------------------------------------------------------------------------
## Superrow var 1
-----------------------------------------------------------------------------*/
foreach srow1 in `srow1_levels' {

	local psrow1: list posof "`srow1'" in srow1_levels
	matrix xt_results_`psrow1' = J(1,`ncol'*`nscol', .)

	if `nsrow1' > 1 {
		tempfile stats_data_`psrow1'
		qui save `stats_data_`psrow1''
		qui keep if `srow1var' == `srow1'
	}


	/*-------------------------------------------------------------------------
	## Superrow var 2
	-------------------------------------------------------------------------*/
	foreach srow2 in `srow2_levels' {

		local psrow2: list posof "`srow2'" in srow2_levels
		matrix xt_results_`psrow1'`psrow2' = J(1,`ncol'*`nscol', .)

		if `nsrow2' > 1 {
			tempfile stats_data_`psrow1'`psrow2'
			qui save `stats_data_`psrow1'`psrow2''
			qui keep if `srow2var' == `srow2'
		}

		/*---------------------------------------------------------------------
		## Superrow var 3
		---------------------------------------------------------------------*/
		foreach srow3 in `srow3_levels' {

			local psrow3: list posof "`srow3'" in srow3_levels
			matrix xt_results_`psrow1'`psrow2'`psrow3' = J(1,`ncol'*`nscol', .)

			if `nsrow3' > 1 {
				tempfile stats_data_`psrow1'`psrow2'`psrow3'
				qui save `stats_data_`psrow1'`psrow2'`psrow3''
				qui keep if `srow3var' == `srow3'
			}

			/*-----------------------------------------------------------------
			## Superrow var 4
			-----------------------------------------------------------------*/
			foreach srow4 in `srow4_levels' {

				local psrow4: list posof "`srow4'" in srow4_levels
				if !missing("`colvar'") {
					mat def xt_results_`psrow1'`psrow2'`psrow3'`psrow4' = ///
														J(`nrow'*`nstats', 1, .)
				}
				else {
					mat def xt_results_`psrow1'`psrow2'`psrow3'`psrow4' = ///
														J(`nrow', 1, .)
				}
				

				if `nsrow4' > 1 {
					tempfile stats_data_`psrow1'`psrow2'`psrow3'`psrow4'
					qui save `stats_data_`psrow1'`psrow2'`psrow3'`psrow4''
					qui keep if `srow4var' == `srow4'
				}


				/*-------------------------------------------------------------
				## Supercolumn
				-------------------------------------------------------------*/
				foreach scol in `scol_levels' {

					if `nscol' > 1 {
						tempfile stats_data_scol
						qui save `stats_data_scol'
						qui keep if `scolvar' == `scol'
					}

					local pscol: list posof "`scol'" in scol_levels
					sort `rowvar' `colvar' `scolvar' `srowvar1' `srowvar2' `srowvar3' `srowvar4'

					if !missing("`colvar'") {
						mat def xt_`pscol' = J(`nrow'*`nstats', `ncol', .)
					}
					else {
						mat def xt_`pscol' = J(`nrow', `nstats', .)
					}

					forvalues row = 1/`nrow' {
						forvalues col = 1/`ncol' {
							forvalues stat = 1/`nstats' {

								if !missing("`colvar'") {
									mat xt_`pscol'[((`row'-1)*`nstats')+`stat', `col'] = ///
												table`stat'[((`row'-1)*`ncol')+`col']
								}
								else {
									mat xt_`pscol'[`row', `col'] = table`col'[`row']
								}

							}
						}
					}


					mat xt_results_`psrow1'`psrow2'`psrow3'`psrow4'  = ///
						xt_results_`psrow1'`psrow2'`psrow3'`psrow4',  xt_`pscol'

					mat drop xt_`pscol'

					if `nscol' > 1 {
						qui use `stats_data_scol', clear
					}

				}
				/*-----------------------------------------------------------*/

				mat xt_results_`psrow1'`psrow2'`psrow3'`psrow4'  = ///
					xt_results_`psrow1'`psrow2'`psrow3'`psrow4'[1..., 2...]

				if `nby' > 0 {
					mat srow_header = J(`nby',`ncol'*`nscol', .)

					mat xt_results_`psrow1'`psrow2'`psrow3' =     ///
					xt_results_`psrow1'`psrow2'`psrow3' \     ///
					srow_header \						       ///
					xt_results_`psrow1'`psrow2'`psrow3'`psrow4' 

					mat drop xt_results_`psrow1'`psrow2'`psrow3'`psrow4' srow_header
				}
				else {
					mat xt_results_`psrow1'`psrow2'`psrow3' =     ///
					xt_results_`psrow1'`psrow2'`psrow3' \     ///
					xt_results_`psrow1'`psrow2'`psrow3'`psrow4' 

					mat drop xt_results_`psrow1'`psrow2'`psrow3'`psrow4' 
				}
				

				
				

				if `nsrow4' > 1 {
					qui use `stats_data_`psrow2'`psrow3'`psrow4'', clear
				}

			}
			/*---------------------------------------------------------------*/


			mat xt_results_`psrow1'`psrow2'`psrow3' = ///
				xt_results_`psrow1'`psrow2'`psrow3'[2..., 1...]

			mat xt_results_`psrow1'`psrow2' = ///
				xt_results_`psrow1'`psrow2' \ xt_results_`psrow1'`psrow2'`psrow3' 

			mat drop xt_results_`psrow1'`psrow2'`psrow3' 

			if `nsrow3' > 1 {
				qui use `stats_data_`psrow1'`psrow2'`psrow3'', clear
			}

		}
		/*-------------------------------------------------------------------*/


		matrix xt_results_`psrow1'`psrow2' = xt_results_`psrow1'`psrow2'[2..., 1...]
		mat xt_results_`psrow1' = xt_results_`psrow1' \ xt_results_`psrow1'`psrow2'
		mat drop xt_results_`psrow1'`psrow2'

		if `nsrow2' > 1 {
			qui use `stats_data_`psrow1'`psrow2'', clear
		}
		
	}
	/*-----------------------------------------------------------------------*/


	matrix xt_results_`psrow1' = xt_results_`psrow1'[2..., 1...]
	mat xt_results = xt_results \ xt_results_`psrow1'
	mat drop xt_results_`psrow1'

	if `nsrow1' > 1 {
		qui use `stats_data_`psrow1'', clear
	}
	
}
/*---------------------------------------------------------------------------*/

mat xtable = xt_results[2..., 1...]
mat drop xt_results


/*********************************************************************
# Labels
**********************************************************************/

/* Rows ans superrows */
foreach row in `row_levels' {
		local row_label : label (`rowvar') `row'
		if `row'== . {
				local row_label "Total"
		}

		#delimit ;
		local mat_rownames_rows = 					
							`"`mat_rownames_rows'"'  
							 + `" ""' 											
							 + subinstr(substr("`row_label'", 1, 30), ".", " ", .) 
							 + `"""' 											
							 					
		;
		#delimit cr

		if !missing("`colvar'"){
			local mat_rownames_rows = `"`mat_rownames_rows'"' + (`""-""')*(`nstats'-1)	
		}

}

foreach srow1 in `srow1_levels' {
	local psrow1: list posof "`srow1'" in srow1_levels
	cap local srow1_label : label (`srow1var') `srow1'
	
	if missing("`srow1_label'"){

		local mat_rownames = `"`mat_rownames_rows'"' 

	} 
	else {
		foreach srow2 in `srow2_levels' {
			local psrow2: list posof "`srow2'" in srow2_levels
			cap local srow2_label : label (`srow2var') `srow2'

			if missing("`srow2_label'"){
				#delimit ;
				local mat_rownames = `"`mat_rownames'"' 
								    + `" ""' 
									+ subinstr(substr("`srow1_label'", 1, 30), ".", " ", .)
							        + `"""'  
							        + `"`mat_rownames_rows'"' 
				;
				#delimit cr
			} 
			else{
				foreach srow3 in `srow3_levels' {
					local psrow3: list posof "`srow3'" in srow3_levels
					cap local srow3_label : label (`srow3var') `srow3'

					if missing("`srow3_label'"){
						#delimit ;
						local mat_rownames = `"`mat_rownames'"' 
										    + `" ""' 
											+ subinstr(substr("`srow1_label'", 1, 30), ".", " ", .)
									        + `"""'  
									        + `" ""' 
											+ subinstr(substr("`srow2_label'", 1, 30), ".", " ", .)
									        + `"""'  
									        + `"`mat_rownames_rows'"' 
						;
						#delimit cr
					} 
					else{
						foreach srow4 in `srow4_levels' {
							local psrow4: list posof "`srow4'" in srow4_levels
							cap local srow4_label : label (`srow4var') `srow4'

							if missing("`srow4_label'"){
								#delimit ;
								local mat_rownames = `"`mat_rownames'"' 
												    + `" ""' 
													+ subinstr(substr("`srow1_label'", 1, 30), ".", " ", .)
											        + `"""'  
											        + `" ""' 
													+ subinstr(substr("`srow2_label'", 1, 30), ".", " ", .)
											        + `"""'  
											        + `" ""' 
													+ subinstr(substr("`srow3_label'", 1, 30), ".", " ", .)
													+ `"""'  
											        + `"`mat_rownames_rows'"' 
								;
								#delimit cr
							}
							else {
								#delimit ;
								local mat_rownames = `"`mat_rownames'"' 
												    + `" ""' 
													+ subinstr(substr("`srow1_label'", 1, 30), ".", " ", .)
											        + `"""'  
											        + `" ""' 
													+ subinstr(substr("`srow2_label'", 1, 30), ".", " ", .)
											        + `"""'  
											        + `" ""' 
													+ subinstr(substr("`srow3_label'", 1, 30), ".", " ", .)
													+ `"""'  
												    + `" ""' 
													+ subinstr(substr("`srow4_label'", 1, 30), ".", " ", .)
													+ `"""'  
											        + `"`mat_rownames_rows'"' 
								;
								#delimit cr

							}
						}
					}					
				}
			}
		}
	}
}

mat rownames xtable = `mat_rownames'


/* Columns and supercolumns */
if !missing("`colvar'") {

	foreach scol in `scol_levels' {

		local pscol: list posof "`scol'" in scol_levels

		foreach col in `col_levels' {
			local col_label : label (`colvar') `col'
			if `col'== . {
				local col_label "Total"
			}
			#delimit ;
			local mat_colnames_`pscol' = `"`mat_colnames_`pscol''"' + `" ""' + 
										subinstr(substr("`col_label'", 1, 30), ".", " ", .) + `"""'
			;
			#delimit cr

		}

		local mat_colnames = `"`mat_colnames'"' + `"`mat_colnames_`pscol''"' 
	}
}
else {

	forvalues s =1/`nstats' {
		local stat_label: var label table`s'
		local mat_colnames = `"`mat_colnames'"' + `" ""' + ///
									subinstr(substr("`stat_label'", 1, 30), ".", " ", .) + `"""'
	}
}
mat colnames xtable = `mat_colnames'


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

return matrix xtable = xtable

di as smcl "Output written to {browse  "`"xtable.xlsx}"'" 
end
