using JuMP
using DataFrames
using CPLEX
using Plots
using StatsPlots
using XLSX

#Rufe die Excelliste "MeritOrder_Excel" und das entsprechende Tabellenblatt ab. Der Datentyp der Tabellenblatt-Inhalte wird ebenfalls definiert
Kategorien = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "Kategorien", infer_eltypes=true)...)
Kraftwerke = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "Kraftwerke", infer_eltypes=true)...)
Energieträger = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "Energieträger", infer_eltypes=true)...)
Nachfrage = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "Nachfrage")...) .|> float
CO2_Preis = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "CO2-Preis")...) .|> float
Verfügbarkeit = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "Verfügbarkeit", infer_eltypes=true)...)


function powerplants(
    powerplant::String,
    efficiency::Float64,
    fuel::String,
)
    return (
        powerplant = powerplant,
        efficiency = efficiency,
        fuel = fuel,
    )
end

n_kraft, m_kraft = size(Kraftwerke)
powerplant_pp = []
for g in 1:n_kraft
    powerplant_g = [
    powerplants(Kraftwerke[g,:Kategorie],Kraftwerke[g,:Wirkungsgrad],Kraftwerke[g,:Energieträger])]
    push!(powerplant_pp, powerplant_g )
end

powerplant_pp
Energieträger

function GK(i)

    f = Brennstoffe[i]
    η = Wirkungsgrade[i]
    p_f = Energieträger[1, f] #Preis je Brennstoff und Brennstoff hängt über f von Kraftwerkskategorie ab
    e_f = Energieträger[2, f] #Emissionsfaktor des Brennstoffes
    p_e = CO2_Preis[1, 1] #CO2-Preis
    #a_f = Verfügbarkeit[:, f] #Gebe die ganze Spalte je nach Brennstoffart aus

    p_el = (p_f / η) + (e_f / η) * p_e  #p_el = Grenzkosten

    return p_el
end

costs=[]

for i in k
    GK(i)
    p_el = GK(i)
    push!(costs, p_el)
end


