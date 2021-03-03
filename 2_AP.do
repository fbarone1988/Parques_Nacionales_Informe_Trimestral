version 14.0
set linesize 255
set more off
clear
set maxvar 32767, perm


use Base\2017_2019-Visit_Ficticio, clear

drop if inlist(AP, "Patagonia", "Pizarro")

keep if inlist(Año, $anio , $anio_ant1)
keep AP Año Mes total
order AP Año Mes total

cap drop if $trimestre == 1 & inlist(Mes,4,5,6,7,8,9,10,11,12)
cap drop if $trimestre == 2 & inlist(Mes,7,8,9,10,11,12)
cap drop if $trimestre == 3 & inlist(Mes,10,11,12)

tempfile Base

replace AP = "B. Petrif." if AP =="Bosques Petrificados"
replace AP = "C. Pantanos" if AP =="Ciervo de los Pantanos"
replace AP = "C. Benítez" if AP =="Colonia Benítez"
replace AP = "Impenetrable" if AP =="El Impenetrable"
replace AP = "L. Blanca" if AP =="Laguna Blanca"
replace AP = "L. Pozuelos" if AP =="Laguna de los Pozuelos"
replace AP = "Glaciares" if AP =="Los Glaciares"
replace AP = "N. Huapi" if AP =="Nahuel Huapi"
replace AP = "N. Toldos" if AP =="Nogalar de los Toldos"
replace AP = "Q. Condorito" if AP =="Quebrada del Condorito"
replace AP = "R. Pilcomayo" if AP =="Río Pilcomayo"
replace AP = "P. Moreno" if AP == "Perito Moreno"
replace AP = "S. Guill." if AP == "San Guillermo"
replace AP = "S. Quijadas" if AP =="Sierra de las Quijadas"
replace AP = "T. Fuego" if AP =="Tierra del Fuego"

save `Base'

gen Trimestre = .
replace Trimestre = 1 if inlist(Mes,1,2,3)
replace Trimestre = 2 if inlist(Mes,4,5,6)
replace Trimestre = 3 if inlist(Mes,7,8,9)
replace Trimestre = 4 if inlist(Mes,10,11,12)

collapse (sum) total, by(AP Año Trimestre)
set obs `=_N+2'

replace Trimestre = 0 if Trimestre == .
replace Año = $anio_ant1 in `=_N-1'
replace Año = $anio in `=_N' 
sort Año Trimestre

qui forval i = $anio_ant1 / $anio {
	summ total if Año == `i'
	replace total = r(sum) if Trimestre == 0 & Año == `i'
}


replace AP = "Total" if Trimestre == 0 

tostring Año, replace
gen Periodo = Año+"_"+string(Trimestre)
drop Trimestre Año
reshape wide total, i(AP) j(Periodo) string 

sort AP
gen orden = _n

replace orden = 0 in L
sort orden
drop orden *_0

ds total*

qui foreach i in `r(varlist)'  {
	summ `i'
	replace `i' = r(sum) in 1
}

local anio = $anio
local anio_ant1 =  $anio_ant1

cap forval i = 1/4 {
	gen var_`i' = total`anio'_`i'/total`anio_ant1'_`i'-1
	label var var_`i' " "
}

cap order AP *1 *2 *3 *4

qui forval i = $anio_ant1/$anio {
	egen total`i' = rowtotal(total`i'*)
	format total`i' %11,0gc
}

gen var_total = total$anio / total$anio_ant1 - 1
label var var_total " "

local trimestre = $trimestre

cap forval i = $anio_ant1 / $anio {
	label var total`i' "`i'*"

	forval j = 1/$trimestre {
		label var total`i'_`j' "`i'*"
	}
}

cap  export excel AP total*_1 var_1                               using           ///
"`anio'/Trimestre `trimestre'/Tablas/T`trimestre'.xlsx", sheet("AP") sheetmodify  ///
keepcellfmt missing("///") cell(A3) firstrow(varlabels) 

cap export excel AP total*_1 var_1 total*_2 var_2                using            ///
"`anio'/Trimestre `trimestre'/Tablas/T`trimestre'.xlsx", sheet("AP") sheetmodify  ///
keepcellfmt missing("///") cell(A3) firstrow(varlabels) 

cap export excel AP total*_1 var_1 total*_2 var_2 total*_3 var_3 using            ///
"`anio'/Trimestre `trimestre'/Tablas/T`trimestre'.xlsx", sheet("AP") sheetmodify  ///
keepcellfmt missing("///") cell(A3) firstrow(varlabels) 

cap export excel AP total*_4 var_4 total$anio_ant1 total$anio var_total using     ///
"`anio'/Trimestre `trimestre'/Tablas/T`trimestre'.xlsx", sheet("AP") sheetmodify  ///
keepcellfmt missing("///") cell(A44) firstrow(varlabels) 

use `Base', clear

collapse (sum) total, by(AP Año Mes)
set obs `=_N+2'

replace Mes = 0 if Mes == .
replace Año = $anio_ant1 in `=_N-1'
replace Año = $anio in `=_N' 
sort Año Mes

qui forval i = $anio_ant1 / $anio {
	summ total if Año == `i'
	replace total = r(sum) if Mes == 0 & Año == `i'
}

replace AP = "Total" if Mes == 0 
tostring Año, replace
gen Periodo = Año+"_"+string(Mes)
drop Mes Año
reshape wide total, i(AP) j(Periodo) string 

gen orden = _n
replace orden = 0 in L
sort orden
drop orden *_0

cap order AP *1 *2 *3, seq
cap order AP *1 *2 *3 *4 *5 *6, seq
cap order AP *1 *2 *3 *4 *5 *6 *7 *8 *9, seq
cap order AP *1 *2 *3 *4 *5 *6 *7 *8 *9 *10 *11 *12, seq

ds total*

qui foreach i in `r(varlist)'  {
	summ `i'
	replace `i' = r(sum) in 1
}

cap forval i = 1/12 {
	gen var_`i' = total`anio'_`i'/total`anio_ant1'_`i'-1
	label var var_`i' " "
	order total`anio_ant1'_`i', before(total`anio'_`i')
	order var_`i', after(total`anio'_`i')
}

qui forval i = $anio_ant1 / $anio {
	egen total`i' = rowtotal(total`i'*)
	format total`i' %11,0gc
}

gen var_total = total$anio / total$anio_ant1 -1
label var var_total " "

if $trimestre == 1 {
	local tope = 3
}

if $trimestre == 2 {
	local tope = 6
}

if $trimestre == 3 {
	local tope = 9
}

if $trimestre == 4 {
	local tope = 12
}

cap forval i = $anio_ant1 / $anio {
	label var total`i' "`i'*"

	forval j = 1/`tope'  {
		label var total`i'_`j' "`i'*"
	}
}

cap export excel AP total*_1 var_1 total*_2 var_2 total*_3 var_3 using            ///
"`anio'/Trimestre `trimestre'/Tablas/T`trimestre'.xlsx", sheet("AP") sheetmodify  ///
keepcellfmt missing("///") cell(L3) firstrow(varlabels) 

cap export excel AP total*_4 var_4 total*_5 var_5 total*_6 var_6 using            ///
"`anio'/Trimestre `trimestre'/Tablas/T`trimestre'.xlsx", sheet("AP") sheetmodify  ///
keepcellfmt missing("///") cell(W3) firstrow(varlabels) 

cap export excel AP total*_7 var_7 total*_8 var_8 total*_9 var_9 using            ///
"`anio'/Trimestre `trimestre'/Tablas/T`trimestre'.xlsx", sheet("AP") sheetmodify  ///
keepcellfmt missing("///") cell(AH3) firstrow(varlabels) 

cap export excel AP total*_10 var_10 total*_11 var_11 total*_12 var_12            ///
total`anio_ant1' total`anio' var_total using                                      ///
"`anio'/Trimestre `trimestre'/Tablas/T`trimestre'.xlsx", sheet("AP") sheetmodify  ///
keepcellfmt missing("///") cell(AS3) firstrow(varlabels) 

exit