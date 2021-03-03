version 14
set linesize 255
set more off
clear

import delimited "Input\2017_2019_Ficticio.csv", case(preserve) encoding(UTF-8)   ///
numericcols(5) clear 

ds Acceso-nactotal
local foo = 1

foreach i in `r(varlist)' {
	local etiq: word `foo' of Cobro "No Residentes" "Nacionales Mayores"          ///
    "Residentes Provinciales" "Residentes Locales" "Estudiantes"                  ///
     "Jubilados y Pensionados" Menores Total Residentes
	label var `i' "`etiq'"
	local foo = `foo'+1
}

compress
cap mkdir Base
save Base/2017_2019-Visit_Ficticio.dta

exit
