version 14.0
set linesize 255
set more off
clear
set maxvar 32767, perm

use Base\2017_2019-Visit_Ficticio, clear
keep if inlist(Año, $anio , $anio_ant1 , $anio_ant2, $anio_ant3 )
keep Año Mes ext nactotal
order Año Mes ext nactotal

cap drop if $trimestre == 1 & inlist(Mes,4,5,6,7,8,9,10,11,12)
cap drop if $trimestre == 2 & inlist(Mes,7,8,9,10,11,12)
cap drop if $trimestre == 3 & inlist(Mes,10,11,12)

gen Trimestre = .
replace Trimestre = 1 if inlist(Mes,1,2,3)
replace Trimestre = 2 if inlist(Mes,4,5,6)
replace Trimestre = 3 if inlist(Mes,7,8,9)
replace Trimestre = 4 if inlist(Mes,10,11,12)

tempfile Base
save `Base'

collapse (sum) ext nactotal, by(Año Mes Trimestre)
set obs `=_N+4'

replace Mes = 0 if Mes == .
replace Año = $anio_ant3 in `=_N-3'
replace Año = $anio_ant2 in `=_N-2'
replace Año = $anio_ant1 in `=_N-1'
replace Año = $anio in `=_N' 
sort Año Mes

qui foreach i in ext nactotal {
	forval j = $anio_ant3 / $anio{
		summ `i' if Año == `j'
		replace `i' = r(sum) if Mes == 0 & Año == `j'
	}
}

qui foreach i in ext nactotal {
		gen var_`i' = .
		gen acum_`i' = . 
		gen var_acum_`i' = .
		bysort Año: replace acum_`i' = sum(`i') if Mes !=0
}

count
local foo = 0
qui foreach i in ext nactotal {
		forval j = `=_N'(-1)1 {
			replace var_`i'= `i'[`j']/`i'[`j'- r(N)/4]-1 in `j'
			replace var_acum_`i'= acum_`i'[`j']/acum_`i'[`j'- r(N)/4]-1 in `j'
			format acum_`i' %11,0gc
		}
local foo = `foo' + 1
}

clonevar Mes_str = Mes
replace Mes_str = Mes + 1
tostring Mes_str, replace
local Mes_str Total Enero Febrero Marzo Abril Mayo Junio Julio Agosto Septiembre  ///
Octubre Noviembre Diciembre

count
local foo = 0 
qui forval i = $anio_ant3 / $anio {
		forval j = 1/13 {
			local Mes_word: word `j' of `Mes_str'
			replace Mes_str = "`Mes_word'" if Año == `i' & Mes_str == "`j'"
		}
	local foo = `foo' + 1
}

label var Año "Año"
label var Mes_str "Mes"
label var nactotal "Visitantes Residentes"
label var ext "Visitantes No Residentes "
qui foreach i in var_nactotal var_ext {
	label var `i' "Variación interanual (%)"
}
order Año Mes *nactotal *ext

keep if inlist(Año, $anio , $anio_ant1)
cap drop if $trimestre == 1 & Trimestre == 0

order Año Mes_str *nactotal *ext

local anio = $anio
local trimestre = $trimestre

preserve
tostring Año, replace
replace Año = " " if Mes_str != "Total"
replace Año = Año+"*" if Mes_str == "Total"

export excel Año Mes_str *nactotal *ext using                                     /// 
"`anio'/Trimestre `trimestre'/Tablas/T`trimestre'.xlsx", sheet("No Residentes")   ///
sheetmodify keepcellfmt missing("///") cell(A3)

restore

global n_graph = $n_graph + 1
graph bar (sum) nactotal ext if Mes != 0, stack over(Año, label(labsize(1.8)       ///
      angle(45)) ) over(Mes_str, sort(Mes ascending) label(labsize(1.5)            ///
      labcolor(59 56 56))) blabel(bar, color(59 56 56) position(center)            ///
      size(vsmall) format(%12,1gc) justification(center))                          ///
	  ylab(,format(%12,1gc) labsize(vsmall) labcolor(59 56 56) nogrid)             ///                                                      
	  title("Gráfico $n_graph.         Visitación mensual: $anio_ant1 - $anio",    ///
	  position(11) color(59 56 56))                                                ///
	  subtitle("                                     Residentes - No Residentes"   ///
      , margin(b=5) position(11) color(59 56 56)) ytitle("Visitantes", size(small) ///
      color(59 56 56)) yscale(lcolor(59 56 56)) ysc(titlegap(5))                   ///
	  graphregion(color(white)) legend(order(1 "Residentes" 2 "No Residentes")     ///
	  region(lwidth(none)) size(vsmall) color(59 56 56))                           ///
	  bar(1, fcolor(eltblue) lcolor(eltblue)) bar(2, fcolor(255 94 174)            ///
      lcolor(255 94 174))                                                          ///
	  note("{bf:Fuente:} DATOS FICTICIOS - Dirección de Mercadeo - Dirección Nacional de Uso Público." "              Administración de Parques Nacionales.", color(59 56 56) size(vsmall))
graph export "`anio'/Trimestre `trimestre'/Gráficos/3_No Residentes/`anio'_`trimestre'_No Residentes_Mensual.png", width(2000) replace


global n_graph = $n_graph + 1
graph bar (sum) nactotal ext, stack percent over(Año,                              ///
      label(labsize(1.8) angle(45)) ) over(Mes_str, sort(Mes ascending)            ///
      label(labsize(1.5) labcolor(59 56 56))) blabel(bar, color(59 56 56)          ///
      position(center) size(vsmall) format(%9.0f) justification(center))           ///
	  ylab(,format(%12,1gc) labsize(vsmall) labcolor(59 56 56) nogrid)             ///                                                      
	  title("Gráfico $n_graph.         Visitación mensual: $anio_ant1 - $anio",    ///
	  position(11) color(59 56 56))                                                ///
	  subtitle("                                     Porcentaje de Residentes - No Residentes"   ///
      , margin(b=5) position(11) color(59 56 56)) ytitle("Visitantes", size(small) ///
      color(59 56 56)) yscale(lcolor(59 56 56)) ysc(titlegap(5))                   ///
	  graphregion(color(white)) legend(order(1 "Residentes" 2 "No Residentes")     ///
	  region(lwidth(none)) size(vsmall) color(59 56 56))                           ///
	  bar(1, fcolor(eltblue) lcolor(eltblue)) bar(2, fcolor(255 94 174)            ///
      lcolor(255 94 174))                                                          ///
	  note("{bf:Fuente:} DATOS FICTICIOS - Dirección de Mercadeo - Dirección Nacional de Uso Público." "              Administración de Parques Nacionales.", color(59 56 56) size(vsmall))
local nb=`.Graph.plotregion1.barlabels.arrnels'
qui forval i=1/`nb' {
		di "`.Graph.plotregion1.barlabels[`i'].text[1]'"
		.Graph.plotregion1.barlabels[`i'].text[1]="`.Graph.plotregion1.barlabels[`i'].text[1]'%"
}
.Graph.drawgraph
graph export "`anio'/Trimestre `trimestre'/Gráficos/3_No Residentes/`anio'_`trimestre'_No Residentes_Mensual_%.png", width(2000) replace


clear
use `Base'
collapse (sum) nactotal *ext, by(Año Trimestre)
set obs `=_N+4'

replace Trimestre = 0 if Trimestre == .
replace Año = $anio_ant3 in `=_N-3'
replace Año = $anio_ant2 in `=_N-2'
replace Año = $anio_ant1 in `=_N-1'
replace Año = $anio in `=_N' 
sort Año Trimestre

qui foreach i in nactotal ext {
	forval j = $anio_ant3 / $anio {
		summ `i' if Año == `j'
		replace `i' = r(sum) if Trimestre == 0 & Año == `j'
	}
}

qui foreach i in ext nactotal {
		gen var_`i' = .
		gen acum_`i' = . 
		gen var_acum_`i' = .
		bysort Año: replace acum_`i' = sum(`i') if Trimestre !=0
}

count
local foo = 0
qui foreach i in ext nactotal {
		forval j = `=_N'(-1)1 {
			replace var_`i'= `i'[`j']/`i'[`j'- r(N)/4]-1 in `j'
			replace var_acum_`i'= acum_`i'[`j']/acum_`i'[`j'- r(N)/4]-1 in `j'
			format acum_`i' %11,0gc
		}
local foo = `foo' + 1
}

clonevar Trimestre_str = Trimestre
replace Trimestre_str = Trimestre + 1
tostring Trimestre_str, replace
local Trimestre_str Total "1er trimestre" "2do trimestre" "3er trimestre"         ///
"4to trimestre"

local foo = 0 
qui forval i = $anio_ant3 / $anio {
		forval j = 1/5 {
			local Trimestre_word: word `j' of `Trimestre_str'
			replace Trimestre_str = "`Trimestre_word'" if Año == `i' &            ///
			Trimestre_str == "`j'"
		}
	local foo = `foo' + 1
}

label var Año "Año"
label var Trimestre_str "Trimestre"
label var nactotal "Visitantes Residentes"
label var ext "Visitantes No Residentes "
qui foreach i in nactotal ext {
	label var var_`i' "Variación interanual (%)"
	label var var_acum_`i' "Variación interanual (%)"
	label var acum_`i' "Acumulado"
}

order Año Trimestre *nactotal *ext

keep if inlist(Año, $anio , $anio_ant1)

cap drop if $trimestre == 1 & Trimestre == 0

preserve
tostring Año, replace

replace Año = " " if Trimestre_str !="Total" & $trimestre !=1
replace Año = Año+"*" if Trimestre_str =="Total" | $trimestre == 1  

export excel Año Trimestre_str *nactotal *ext using                               /// 
"`anio'/Trimestre `trimestre'/Tablas/T`trimestre'.xlsx", sheet("No Residentes")   ///
sheetmodify keepcellfmt missing("///") cell(L3)

restore

global n_graph = $n_graph + 1
graph bar (sum) nactotal ext if Trimestre != 0, stack over(Año, label(labsize(1.8) ///
      angle(45)) ) over(Trimestre_str, sort(Mes ascending) label(labsize(1.5)      ///
      labcolor(59 56 56))) blabel(bar, color(59 56 56) position(center)            ///
      size(vsmall) format(%12,1gc) justification(center))                            ///
	  ylab(,format(%12,1gc) labsize(vsmall) labcolor(59 56 56) nogrid)             ///                                                      
	  title("Gráfico $n_graph.         Visitación trimestral: $anio_ant1 - $anio", ///
	  position(11) color(59 56 56))                                                ///
	  subtitle("                                     Residentes - No Residentes"   ///
      , margin(b=5) position(11) color(59 56 56)) ytitle("Visitantes", size(small) ///
      color(59 56 56)) yscale(lcolor(59 56 56)) ysc(titlegap(5))                   ///
	  graphregion(color(white)) legend(order(1 "Residentes" 2 "No Residentes")     ///
	  region(lwidth(none)) size(vsmall) color(59 56 56))                           ///
	  bar(1, fcolor(eltblue) lcolor(eltblue)) bar(2, fcolor(255 94 174)            ///
      lcolor(255 94 174))                                                          ///
	  note("{bf:Fuente:} DATOS FICTICIOS - Dirección de Mercadeo - Dirección Nacional de Uso Público." "              Administración de Parques Nacionales.", color(59 56 56) size(vsmall))
graph export "`anio'/Trimestre `trimestre'/Gráficos/3_No Residentes/`anio'_`trimestre'_No Residentes_Trimestral.png", width(2000) replace

global n_graph = $n_graph + 1
graph bar (sum) nactotal ext, stack percent over(Año,                               ///
      label(labsize(1.8) angle(45)) ) over(Trimestre_str, sort(Trimestre ascending) ///
      label(labsize(1.5) labcolor(59 56 56))) blabel(bar, color(59 56 56)          ///
      position(center) size(vsmall) format(%9.0f) justification(center))         ///
	  ylab(,format(%12,1gc) labsize(vsmall) labcolor(59 56 56) nogrid)             ///                                                      
	  title("Gráfico $n_graph.         Visitación trimestral: $anio_ant1 - $anio", ///
	  position(11) color(59 56 56))                                                ///
	  subtitle("                                     Porcentaje de Residentes - No Residentes"   ///
      , margin(b=5) position(11) color(59 56 56)) ytitle("Visitantes", size(small) ///
      color(59 56 56)) yscale(lcolor(59 56 56)) ysc(titlegap(5))                   ///
	  graphregion(color(white)) legend(order(1 "Residentes" 2 "No Residentes")     ///
	  region(lwidth(none)) size(vsmall) color(59 56 56))                           ///
	  bar(1, fcolor(eltblue) lcolor(eltblue)) bar(2, fcolor(255 94 174)            ///
      lcolor(255 94 174))                                                          ///
	  note("{bf:Fuente:} DATOS FICTICIOS - Dirección de Mercadeo - Dirección Nacional de Uso Público." "              Administración de Parques Nacionales.", color(59 56 56) size(vsmall))
local nb=`.Graph.plotregion1.barlabels.arrnels'
qui forval i=1/`nb' {
		di "`.Graph.plotregion1.barlabels[`i'].text[1]'"
		.Graph.plotregion1.barlabels[`i'].text[1]="`.Graph.plotregion1.barlabels[`i'].text[1]'%"
}
.Graph.drawgraph
graph export "`anio'/Trimestre `trimestre'/Gráficos/3_No Residentes/`anio'_`trimestre'_No Residentes_Trimestral_%.png", width(2000) replace

exit
