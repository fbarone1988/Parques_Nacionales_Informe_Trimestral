local anios 2019

local error "El año ingresado es incorrecto."
cls
cap while 1 {
    noi display "Ingresar año del informe."
	noi display in smcl as input                                                  ///
	"Años válidos: `anios'. VERSIÓN GIT: SOLO 2019 HABILITADO.", _request(anio)
    local correct : list global(anio) in anios
    if !`correct' noi display in smcl as error "`error'"
    else continue, break
}
cls

global anio_ant1 = $anio - 1
global anio_ant2 = $anio - 2
global anio_ant3 = $anio - 3

cls
local trimestres 1 2 3 4
local error "El trimestre ingresado es incorrecto."
cls
cap while 1 {
    noi display in smcl as input "Elegir trimestre del informe:", _request(trimestre)
    local correct : list global(trimestre) in trimestres
    if !`correct' noi display in smcl as error "`error'"
    else continue, break
}

cap mkdir "$anio"
cap mkdir "$anio/Trimestre $trimestre"
cap mkdir "$anio/Trimestre $trimestre/Gráficos"
cap mkdir "$anio/Trimestre $trimestre/Gráficos/1_Total"
cap mkdir "$anio/Trimestre $trimestre/Gráficos/3_No Residentes"
cap mkdir "$anio/Trimestre $trimestre/Gráficos/4_Residentes"
cap mkdir "$anio/Trimestre $trimestre/Gráficos/5_Regional"
cap mkdir "$anio/Trimestre $trimestre/Tablas"

cls
di "El trimestre elegido fue el: $trimestre."

local path `c(pwd)'
global n_graph = 0

run "1_Total.do"
qui cd "`path'"

run "2_AP.do"
qui cd "`path'"

run "3_No Residentes.do"
qui cd "`path'"

run "4_Residentes.do"
qui cd "`path'"

run "5_Regional"
qui cd "`path'"
clear

graph close

exit