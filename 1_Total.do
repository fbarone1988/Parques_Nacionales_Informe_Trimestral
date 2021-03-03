version 14.0
set linesize 255
set more off
clear
set maxvar 32767, perm

use Base\2017_2019-Visit_Ficticio, clear

keep if inlist(Año, $anio , $anio_ant1 , $anio_ant2, $anio_ant3 )

keep Año Mes total
order Año Mes total

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

collapse (sum) total, by(Año Mes Trimestre)
set obs `=_N+4'

replace Mes = 0 if Mes == .
replace Año = $anio_ant3 in `=_N-3'
replace Año = $anio_ant2 in `=_N-2'
replace Año = $anio_ant1 in `=_N-1'
replace Año = $anio in `=_N' 
sort Año Mes

qui forval i = $anio_ant3 / $anio {
	summ total if Año == `i'
	replace total = r(sum) if Mes == 0 & Año == `i'
}

gen acum = . 
bysort Año: replace acum = sum(total) if Mes !=0
format acum %11,0gc

count
local foo = 0
qui foreach i in total acum {
		gen var_`i' = .
		forval j = `=_N'(-1)1 {
		replace var_`i'= `i'[`j']/`i'[`j'- r(N)/4]-1 in `j'
		}
	local foo = `foo' + 1
}

keep if inlist(Año, $anio , $anio_ant1)

clonevar Mes_str = Mes
replace Mes_str = Mes + 1
tostring Mes_str, replace
local Mes_str Total Enero Febrero Marzo Abril Mayo Junio Julio Agosto Septiembre     ///
Octubre Noviembre Diciembre

local foo = 0 
qui forval i = $anio_ant1 / $anio  {
		forval j = 1/13 {
			local Mes_word: word `j' of `Mes_str'
			replace Mes_str = "`Mes_word'" if Año == `i' & Mes_str == "`j'"
		}
	local foo = `foo' + 1
}

label var Año "Año"
label var Mes_str "Mes"
label var total "Visitantes"
label var var_total "Variación interanual (%)"
label var acum Acumulado
label var var_acum "Variación interanual (%)"

order Año Mes_str total var_total acum var_acum

//cd "Informe trimestral"

local anio $anio
local trimestre $trimestre

preserve
tostring Año, replace
replace Año = " " if Mes_str !="Total"
replace Año = Año+"*" if Mes_str =="Total" 

export excel Año Mes_str total var_total acum var_acum using                         /// 
"`anio'/Trimestre `trimestre'/Tablas/T`trimestre'.xlsx", sheet("Total") sheetmodify  ///
keepcellfmt missing("///") cell(A3)

restore

keep total Año Mes Mes_str
reshape wide total, i(Mes Mes_str) j(Año)


global n_graph = $n_graph + 1
graph bar (sum) total* if Mes != 0,                                                  ///
	  over(Mes_str, sort(Mes ascending) label(labsize(1.8)                           ///
      labcolor(59 56 56))) blabel(bar, color(59 56 56)                               ///
	  position(outside) size(vsmall) format(%12,1gc) justification(center))          ///
      ylab(,format(%12,1gc) labsize(vsmall) labcolor(59 56 56) nogrid)               ///                                                      
	  title("Gráfico $n_graph.         Visitación mensual: $anio_ant1 - $anio",      ///
      position(11) color(59 56 56)) subtitle(" ", margin(1+1 r+1 b-1 t-1))           ///
	  ytitle("Visitantes", size(small) color(59 56 56))                              ///
	  yscale(lcolor(59 56 56)) ysc(titlegap(5))                                      ///
      graphregion(color(white)) legend(order(1 "$anio_ant1" 2 "$anio")               ///
      region(lwidth(none)) size(vsmall) color(59 56 56))                             ///
	  bar(1, fcolor(gs11) lcolor(gs11)) bar(2, fcolor(edkblue) lcolor(edkblue))      ///
	  note("{bf:Fuente:} DATOS FICTICIOS - Dirección de Mercadeo - Dirección Nacional de Uso Público." "              Administración de Parques Nacionales.", color(59 56 56) size(vsmall))
graph export "`anio'/Trimestre `trimestre'/Gráficos/1_Total/`anio'_`trimestre'_Total_Mensual.png", width(2000) replace


clear
use `Base'
collapse (sum) total, by(Año Trimestre)
set obs `=_N+4'

replace Trimestre = 0 if Trimestre == .
replace Año = $anio_ant3 in `=_N-3'
replace Año = $anio_ant2 in `=_N-2'
replace Año = $anio_ant1 in `=_N-1'
replace Año = $anio in `=_N'  
sort Año Trimestre

qui forval i = $anio_ant3 / $anio  {
		summ total if Año == `i'
		replace total = r(sum) if Trimestre == 0 & Año == `i'
}

gen acum = . 
bysort Año: replace acum = sum(total) if Trimestre !=0
format acum %11,0gc

count
local foo = 0
qui foreach i in total acum {
		gen var_`i' = .
		forval j = `=_N'(-1)1 {
		replace var_`i'= `i'[`j']/`i'[`j'-r(N)/4]-1 in `j'
	}
	local foo = `foo' + 1
}

clonevar Trimestre_str = Trimestre
replace Trimestre_str = Trimestre + 1
tostring Trimestre_str, replace
local Trimestre_str Total "1er trimestre" "2do trimestre" "3er trimestre"            ///
"4to trimestre"

keep if inlist(Año, $anio , $anio_ant1)

local foo = 0 
qui forval i = $anio_ant1 / $anio {
		forval j = 1/5 {
			local Trimestre_word: word `j' of `Trimestre_str'
			replace Trimestre_str = "`Trimestre_word'" if Año == `i' &               ///
			Trimestre_str == "`j'"
		}
	local foo = `foo' + 1
}

label var Año "Año"
label var Trimestre_str "Trimestre"
label var total "Visitantes"
label var var_total "Variación interanual (%)"
label var acum Acumulado
label var var_acum "Variación interanual (%)"

cap drop if $trimestre == 1 & Trimestre == 0

order Año Trimestre_str total var_total acum var_acum

preserve
tostring Año, replace
replace Año = " " if Trimestre_str !="Total" & $trimestre !=1
replace Año = Año+"*" if Trimestre_str =="Total" | $trimestre == 1  

export excel Año Trimestre_str total var_total acum var_acum using                   /// 
"`anio'/Trimestre `trimestre'/Tablas/T`trimestre'.xlsx", sheet("Total")              ///
sheetmodify keepcellfmt missing("///") cell(A33)

restore

keep total Año Trimestre Trimestre_str
reshape wide total, i(Trimestre Trimestre_str) j(Año)

global n_graph = $n_graph + 1

graph bar (sum) total* if Trimestre != 0,                                            ///
	  over(Trimestre_str, sort(Trimestre ascending) label(labsize(small)             ///
      labcolor(59 56 56))) blabel(bar, color(59 56 56)                               ///
	  position(outside) size(vsmall) format(%12,1gc) justification(center))          ///
      ylab(,format(%12,1gc) labsize(vsmall) labcolor(59 56 56) nogrid)               ///                                                      
	  title("Gráfico $n_graph.         Visitación trimestral: $anio_ant1 - $anio",   ///
      position(11) color(59 56 56)) subtitle(" ", margin(1+1 r+1 b-1 t-1))           ///
	  ytitle("Visitantes", size(small) color(59 56 56))                              ///
	  yscale(lcolor(59 56 56)) ysc(titlegap(5))                                      ///
      graphregion(color(white)) legend(order(1 "$anio_ant1" 2 "$anio")               ///
      region(lwidth(none)) size(vsmall) color(59 56 56))                             ///
	  bar(1, fcolor(gs11) lcolor(gs11)) bar(2, fcolor(edkblue) lcolor(edkblue))      ///
	  note("{bf:Fuente:} DATOS FICTICIOS - Dirección de Mercadeo - Dirección Nacional de Uso Público." "              Administración de Parques Nacionales.", color(59 56 56) size(vsmall))
graph export "`anio'/Trimestre `trimestre'/Gráficos/1_Total/`anio'_`trimestre'_Total_Trimestral.png", width(2000) replace

exit
