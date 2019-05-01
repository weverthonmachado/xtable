/*---------------------------------------------------------------------		
/// XTABLE - export table output to excel

/// Weverthon Machado

v0.1 - 2019-05-01
---------------------------------------------------------------------*/
capture program drop xtable
program define xtable
version 13,1

/* Especicando apenas as opções de table que preciso usar como locals.
O asterisco no final da conta de aceitar todas as outras */
syntax varlist(max=3) [if] [in] [fw aw pw iw] [, /* 
		*/ BY(varlist) COLumn REPLACE ROW SColumn  *]


/* Tokenize */
tokenize
local tableargs "`0'"

tokenize `varlist'
local rowvar = "`1'"
local colvar = "`2'"
local scvar = "`3'"
local srvar =  "`by'"


/* Roda table com replace */
tempfile originaldata
qui save `originaldata'
table `tableargs' replace


/* número de tables (ie, o numero de estatisticas calculadas em cont()) */
unab tablelist: table*
local ntables: word count `tablelist'

qui levelsof `rowvar', local(r_levels)
qui levelsof `colvar', local(c_levels)
if !missing("`srvar'") {
	qui levelsof `srvar', local(sr_levels)
}
else {
	local sr_levels 1
}


local nrow: list sizeof r_levels
local ncol: list sizeof c_levels
if !missing("`srvar'") {
	local nsr:  list sizeof sr_levels
} 
else {
	local nsr = 1
}

cap drop _fillin
fillin `rowvar' `colvar' `srvar'
drop _fillin


/*--------------------------------------------------------------------
## Com superrow: uma submatrix (nrow*ntables)*(ncol) pra cada categoria da srow
---------------------------------------------------------------------*/
if !missing("`srvar'") {
	preserve
	/*--------------------------------
	### Labels para submatrices
	----------------------------------*/
	qui levelsof `rowvar', local(varlevels)
	foreach l in `varlevels' {
		local new : label (`1') `l'
		local labels_`rowvar' = `"`labels_`rowvar''"' + `" ""' + subinstr(substr("`new'", 1, 30), ".", " ", .) + `"""'
	}

	/*--------------------------------
	### Cria cada submatrix
	----------------------------------*/
	foreach s in `sr_levels' {
		qui keep if `srvar' == `s'
		sort `rowvar' `colvar'

		mat sr`s' = J(`nrow'*`ntables', `ncol', .)
		local matrix_rownames = (`"`labels_`rowvar''"' + " ") * `ntables'
		mat rownames sr`s' = `matrix_rownames'

		forvalues r = 1/`nrow' {
			forvalues c = 1/`ncol' {
				forvalues t = 1/`ntables' {
					mat sr`s'[((`r'-1)*`ntables')+`t', `c'] = table`t'[((`r'-1)*`ncol')+`c']	
				}
			}
		}
		restore, preserve
	}
	restore

	/*--------------------------------
	### Restaura dados e reporta
	----------------------------------*/
	/* Recupera dados originais */
	qui use `originaldata', replace

	/* Junta matrizes e reporta */
	matrix xtable = J(1,`ncol', .)
	mat colnames xtable = `c_levels'

	foreach s in `sr_levels' {
		matrix head = J(1,`ncol',.)
		matrix rownames head =  `s'
		matrix xtable =  xtable \ head \ sr`s'
		matrix drop sr`s' head 
	}

}

/*--------------------------------------------------------------------
## Sem superrow: uma submatrix (ntables)*(ncol) pra cada categoria da row
---------------------------------------------------------------------*/

if missing("`srvar'") {
	preserve
	/*--------------------------------
	### Labels para submatrices
	----------------------------------*/
	forvalues t = 1/`ntables'  {
		local table_label: variable label table`t'
		/* remove pontos das var labels, para poder usar com mat rownames */
		*local labels_tables = "`labels_tables'" + " " + subinstr("`table_label'" , ".", " ", .)
		local labels_tables = `"`labels_tables'"' + `" ""' + subinstr(substr("`table_label'", 1, 30), ".", " ", .) + `"""'

	}

	/*--------------------------------
	### Cria cada submatrix
	----------------------------------*/
	foreach r in `r_levels' {

		qui keep if `rowvar' == `r'
		sort `colvar'

		mat r`r' = J(`ntables', `ncol', .)
		if `ntables' == 1 {
			matrix rownames r`r' = `r'
		}
		else {
			mat rownames r`r' = `labels_tables'
		}
		

		forvalues c = 1/`ncol' {
			forvalues t = 1/`ntables' {
				mat r`r'[`t', `c'] = table`t'[`c']	
			}
		}
		restore, preserve
	}
	restore

	/*--------------------------------
	### Restaura dados e reporta
	----------------------------------*/
	/* Recupera dados originais */
	qui use `originaldata', replace

	/* Junta matrizes e reporta */
	matrix xtable = J(1,`ncol', .)
	mat colnames xtable = `c_levels'

	foreach r in `r_levels' {
		/* Se for apenas uma estatística, não tem header */
		if `ntables' == 1 {
			matrix xtable =  xtable \ r`r'
		}
		else {
			matrix head = J(1,`ncol',.)
			matrix rownames head =  `r'
			matrix xtable =  xtable \ head \ r`r'
			matrix drop head
		}
		matrix drop r`r'
	}

}


/* Exporta */
/* Remove linha em branco na matriz de resultados */
mat xtable = xtable[2..., 1...]

putexcel set xtable.xlsx, replace
qui putexcel A1 = matrix(xtable, names)

di as smcl "output written to {browse  "`"xtable.xlsx}"'" 
end
