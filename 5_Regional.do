version 14.0
set linesize 255
set more off
clear
set maxvar 32767, perm

use Base\2017_2019-Visit_Ficticio, clear

keep if inlist(Año, $anio , $anio_ant1 , $anio_ant2, $anio_ant3 )
keep Año Mes Región total
order Año Mes total

cap drop if $trimestre == 1 & inlist(Mes,4,5,6,7,8,9,10,11,12)
cap drop if $trimestre == 2 & inlist(Mes,7,8,9,10,11,12)
cap drop if $trimestre == 3 & inlist(Mes,10,11,12)

collapse (sum) total, by(Región Año Mes)

tempfile Base
save `Base'

levelsof Región, local(regiones)

sort Región Año Mes 
reshape wide total , i(Mes Año) j(Región) string 

set obs `=_N+4'
replace Año = $anio_ant3 in `=_N-3'
replace Año = $anio_ant2 in `=_N-2'
replace Año = $anio_ant1 in `=_N-1'
replace Año = $anio in `=_N' 
replace Mes = 0 if Mes == .
sort Año Mes 
 
qui forval i = $anio_ant3 / $anio {
	foreach j of local regiones {  
		summ total`j' if Año == `i'
		replace total`j' = `r(sum)' if Año == `i' & Mes == 0
	}
}

qui foreach i of local regiones {  
	gen acum_`i' = .
	bysort Año: replace acum_`i' = sum(total`i') if Mes !=0
	format acum_`i' %11,0gc
}

count
local foo = 0
qui foreach i of local regiones {  
	gen var_`i' = .
	gen var_acum_`i' = .
	forval j = `=_N'(-1)1 {
	replace var_`i'= total`i'[`j']/total`i'[`j'- r(N)/4]-1 in `j'
	replace var_acum_`i'= acum_`i'[`j']/acum_`i'[`j'- r(N)/4]-1 in `j'
	}
	local foo = `foo' + 1
}

qui foreach i of local regiones {  
	order var_`i', after(total`i')
	order acum_`i', after(var_`i')
	order var_acum_`i', after(acum_`i')
}

local Mes_str Enero Febrero Marzo Abril Mayo Junio Julio Agosto Septiembre        ///
Octubre Noviembre Diciembre

local foo = 1 

gen Mes_str = ""
qui forval i = $anio_ant3 / $anio {
	forval j = 1/12 {
		local Mes_word: word `j' of `Mes_str'
		replace Mes_str = "`Mes_word'" if Año == `i' & Mes == `j'
	}
local foo = `foo' + 1
}
replace Mes_str = "Total" if Mes ==0

label var Año "Año"
label var Mes_str "Mes"

qui foreach i of local regiones {  
	label var total`i' "Visitantes `i'"
	label var var_`i' "Variación interanual (%)"
	label var acum_`i' Acumulado
	label var var_acum_`i' "Variación interanual (%)"
}

keep if inlist(Año, $anio , $anio_ant1)
cap erase "Informe trimestral/5_Regional.xlsx"
local anio = $anio
local trimestre = $trimestre

preserve
tostring Año, replace
replace Año = " " if Mes_str !="Total"
replace Año = Año+"*" if Mes_str == "Total"

export excel Año Mes_str *CENTRO using "`anio'/Trimestre `trimestre'/Tablas/T`trimestre'.xlsx",    ///
sheet("Regional") sheetmodify keepcellfmt missing("///") cell(A4)

export excel Año Mes_str *NEA using "`anio'/Trimestre `trimestre'/Tablas/T`trimestre'.xlsx",       ///
sheet("Regional") sheetmodify keepcellfmt missing("///") cell(A35)

export excel Año Mes_str *NOA using "`anio'/Trimestre `trimestre'/Tablas/T`trimestre'.xlsx",       ///
sheet("Regional") sheetmodify keepcellfmt missing("///") cell(A66)

export excel Año Mes_str *PATAGONIA using "`anio'/Trimestre `trimestre'/Tablas/T`trimestre'.xlsx", ///
sheet("Regional") sheetmodify keepcellfmt missing("///") cell(A97)

restore

local nombres Centro NEA NOA Patagonia

local foo = 1
foreach i in CENTRO NEA NOA PATAGONIA {
	local reemplazo: word `foo' of `nombres'
	rename total`i' `reemplazo'
	local foo = `foo' + 1
}

keep Centro NEA NOA Patagonia Año Mes Mes_str
reshape wide Centro NEA NOA Patagonia, i(Mes Mes_str) j(Año)

foreach i in Centro {
	global n_graph = $n_graph + 1
	graph bar (sum) `i'* if Mes != 0,                                                    ///
		  over(Mes_str, sort(Mes ascending) label(labsize(1.8)                           ///
		  labcolor(59 56 56))) blabel(bar, color(59 56 56)                               ///
		  position(outside) size(vsmall) format(%12,1gc) justification(center))          ///
		  ylab(,format(%12,1gc) labsize(vsmall) labcolor(59 56 56) nogrid)               ///                                                      
		  title("Gráfico $n_graph.         Visitación mensual: $anio_ant1 - $anio",      ///
		  position(11) color(59 56 56)) subtitle("Región `i'", margin(1+1 r+1 b+3 t-1))  ///
		  ytitle("Visitantes", size(small) color(59 56 56))                              ///
		  yscale(lcolor(59 56 56)) ysc(titlegap(5))                                      ///
		  graphregion(color(white)) legend(order(1 "$anio_ant1" 2 "$anio")               ///
		  region(lwidth(none)) size(vsmall) color(59 56 56))                             ///
		  bar(1, fcolor(gs11) lcolor(gs11)) bar(2, fcolor(edkblue) lcolor(edkblue))      ///
		  note("{bf:Fuente:} DATOS FICTICIOS - Dirección de Mercadeo - Dirección Nacional de Uso Público." "              Administración de Parques Nacionales.", color(59 56 56) size(vsmall))
	graph export "`anio'/Trimestre `trimestre'/Gráficos/5_Regional/`anio'_`trimestre'_No Residentes_Mensual_`i'.png", width(2000) replace
}

foreach i in NEA NOA Patagonia {
	global n_graph = $n_graph + 2
	graph bar (sum) `i'* if Mes != 0,                                                    ///
		  over(Mes_str, sort(Mes ascending) label(labsize(1.8)                           ///
		  labcolor(59 56 56))) blabel(bar, color(59 56 56)                               ///
		  position(outside) size(vsmall) format(%12,1gc) justification(center))          ///
		  ylab(,format(%12,1gc) labsize(vsmall) labcolor(59 56 56) nogrid)               ///                                                      
		  title("Gráfico $n_graph.         Visitación mensual: $anio_ant1 - $anio",      ///
		  position(11) color(59 56 56)) subtitle("Región `i'", margin(1+1 r+1 b+3 t-1))  ///
		  ytitle("Visitantes", size(small) color(59 56 56))                              ///
		  yscale(lcolor(59 56 56)) ysc(titlegap(5))                                      ///
		  graphregion(color(white)) legend(order(1 "$anio_ant1" 2 "$anio")               ///
		  region(lwidth(none)) size(vsmall) color(59 56 56))                             ///
		  bar(1, fcolor(gs11) lcolor(gs11)) bar(2, fcolor(edkblue) lcolor(edkblue))      ///
		  note("{bf:Fuente:} DATOS FICTICIOS - Dirección de Mercadeo - Dirección Nacional de Uso Público." "              Administración de Parques Nacionales.", color(59 56 56) size(vsmall))
	graph export "`anio'/Trimestre `trimestre'/Gráficos/5_Regional/`anio'_`trimestre'_No Residentes_Mensual_`i'.png", width(2000) replace
}


use `Base', clear

gen Trimestre = .
replace Trimestre = 1 if inlist(Mes,1,2,3)
replace Trimestre = 2 if inlist(Mes,4,5,6)
replace Trimestre = 3 if inlist(Mes,7,8,9)
replace Trimestre = 4 if inlist(Mes,10,11,12)

collapse (sum) total, by(Región Año Trimestre)

levelsof Región, local(regiones)

sort Región Año Trimestre
reshape wide total , i(Trimestre Año) j(Región) string 

set obs `=_N+4'
replace Año = $anio_ant3 in `=_N-3'
replace Año = $anio_ant2 in `=_N-2'
replace Año = $anio_ant1 in `=_N-1'
replace Año = $anio in `=_N' 
replace Trimestre = 0 if Trimestre == .
sort Año Trimestre
 
qui forval i = $anio_ant3 / $anio {
	foreach j of local regiones {  
		summ total`j' if Año == `i'
		replace total`j' = `r(sum)' if Año == `i' & Trimestre == 0
	}
}

qui foreach i of local regiones {  
	gen acum_`i' = .
	gen var_`i' = .
	gen var_acum_`i' = .
	bysort Año: replace acum_`i' = sum(total`i') if Trimestre !=0
	format acum_`i' %11,0gc
}

count
local foo = 0
qui foreach i of local regiones {  
		forval j = `=_N'(-1)1 {
		replace var_`i'= total`i'[`j']/total`i'[`j'- r(N)/4]-1 in `j'
		replace var_acum_`i'= acum_`i'[`j']/acum_`i'[`j'- r(N)/4]-1 in `j'
	}
		local foo = `foo' + 1
}

qui foreach i of local regiones {  
	order var_`i', after(total`i')
	order acum_`i', after(var_`i')
	order var_acum_`i', after(acum_`i')
}

replace Trimestre = Trimestre +1 
tostring Trimestre, replace
local Trimestre Total "1er trimestre" "2do trimestre" "3er trimestre" "4to trimestre"

local foo = 0 
forval i = $anio_ant3 / $anio {
		forval j = 1/5 {
			local Trimestre_word: word `j' of `Trimestre'
			replace Trimestre = "`Trimestre_word'" if Año == `i' &                ///
			Trimestre == "`j'"
		}
	local foo = `foo' + 1
}

label var Año "Año"
label var Trimestre "Trimestre"

qui foreach i of local regiones {  
	label var total`i' "Visitantes `i'"
	label var var_`i' "Variación interanual (%)"
	label var acum_`i' Acumulado
	label var var_acum_`i' "Variación interanual (%)"
}

keep if inlist(Año, $anio , $anio_ant1)

cap drop if $trimestre == 1 & Trimestre == "Total"

preserve
tostring Año, replace
replace Año = " " if Trimestre !="Total" & $trimestre !=1
replace Año = Año+"*" if Trimestre =="Total" | $trimestre == 1  

export excel Año Trimestre *CENTRO using "`anio'/Trimestre `trimestre'/Tablas/T`trimestre'.xlsx",      ///
sheet("Regional") sheetmodify keepcellfmt missing("///") cell(H4)

export excel Año Trimestre *NEA using "`anio'/Trimestre `trimestre'/Tablas/T`trimestre'.xlsx",         ///
sheet("Regional") sheetmodify keepcellfmt missing("///") cell(H35)

export excel Año Trimestre *NOA using "`anio'/Trimestre `trimestre'/Tablas/T`trimestre'.xlsx",         ///
sheet("Regional") sheetmodify keepcellfmt missing("///") cell(H66)

export excel Año Trimestre *PATAGONIA using "`anio'/Trimestre `trimestre'/Tablas/T`trimestre'.xlsx",   ///
sheet("Regional") sheetmodify keepcellfmt missing("///") cell(H97)

restore

local nombres Centro NEA NOA Patagonia

local foo = 1
foreach i in CENTRO NEA NOA PATAGONIA {
	local reemplazo: word `foo' of `nombres'
	rename total`i' `reemplazo'
	local foo = `foo' + 1
}

keep Centro NEA NOA Patagonia Año Trimestre
reshape wide Centro NEA NOA Patagonia, i(Trimestre) j(Año)

global n_graph = $n_graph - 6

foreach i in Centro {
	global n_graph = $n_graph + 1
	graph bar (sum) `i'* if Trimestre != "Total",                                        ///
		  over(Trimestre, sort(Trimestre ascending) label(labsize(1.8)                   ///
		  labcolor(59 56 56))) blabel(bar, color(59 56 56)                               ///
		  position(outside) size(vsmall) format(%12,1gc) justification(center))          ///
		  ylab(,format(%12,1gc) labsize(vsmall) labcolor(59 56 56) nogrid)               ///                                                      
		  title("Gráfico $n_graph.         Visitación trimestral: $anio_ant1 - $anio",   ///
		  position(11) color(59 56 56)) subtitle("Región `i'", margin(1+1 r+1 b+3 t-1))  ///
		  ytitle("Visitantes", size(small) color(59 56 56))                              ///
		  yscale(lcolor(59 56 56)) ysc(titlegap(5))                                      ///
		  graphregion(color(white)) legend(order(1 "$anio_ant1" 2 "$anio")               ///
		  region(lwidth(none)) size(vsmall) color(59 56 56))                             ///
		  bar(1, fcolor(gs11) lcolor(gs11)) bar(2, fcolor(edkblue) lcolor(edkblue))      ///
		  note("{bf:Fuente:} DATOS FICTICIOS - Dirección de Mercadeo - Dirección Nacional de Uso Público." "              Administración de Parques Nacionales.", color(59 56 56) size(vsmall))
	graph export "`anio'/Trimestre `trimestre'/Gráficos/5_Regional/`anio'_`trimestre'_No Residentes_Trimestral_`i'.png", width(2000) replace
}

foreach i in NEA NOA Patagonia {
	global n_graph = $n_graph + 2
	graph bar (sum) `i'* if Trimestre != "Total",                                        ///
		  over(Trimestre, sort(Trimestre ascending) label(labsize(1.8)                   ///
		  labcolor(59 56 56))) blabel(bar, color(59 56 56)                               ///
		  position(outside) size(vsmall) format(%12,1gc) justification(center))          ///
		  ylab(,format(%12,1gc) labsize(vsmall) labcolor(59 56 56) nogrid)               ///                                                      
		  title("Gráfico $n_graph.         Visitación trimestral: $anio_ant1 - $anio",   ///
		  position(11) color(59 56 56)) subtitle("Región `i'", margin(1+1 r+1 b+3 t-1))  ///
		  ytitle("Visitantes", size(small) color(59 56 56))                              ///
		  yscale(lcolor(59 56 56)) ysc(titlegap(5))                                      ///
		  graphregion(color(white)) legend(order(1 "$anio_ant1" 2 "$anio")               ///
		  region(lwidth(none)) size(vsmall) color(59 56 56))                             ///
		  bar(1, fcolor(gs11) lcolor(gs11)) bar(2, fcolor(edkblue) lcolor(edkblue))      ///
		  note("{bf:Fuente:} DATOS FICTICIOS - Dirección de Mercadeo - Dirección Nacional de Uso Público." "              Administración de Parques Nacionales.", color(59 56 56) size(vsmall))
	graph export "`anio'/Trimestre `trimestre'/Gráficos/5_Regional/`anio'_`trimestre'_No Residentes_Trimestral_`i'.png", width(2000) replace
}

exit

