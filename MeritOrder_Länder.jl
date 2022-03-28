#Extensions
using JuMP
using CPLEX
using XLSX, DataFrames

#Rufe die Excelliste "MeritOrder_Excel" und das entsprechende Tabellenblatt ab. Der Datentyp der Tabellenblatt-Inhalte wird ebenfalls definiert
Kategorien = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "Kategorien", infer_eltypes=true)...)
Kraftwerke = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "Kraftwerke", infer_eltypes=true)...)
Energieträger = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "Energieträger", infer_eltypes=true)...)
Nachfrage = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "Nachfrage")...) .|> float
CO2_Preis = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "CO2-Preis")...) .|> float
Verfügbarkeit = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "Verfügbarkeit", infer_eltypes=true)...)

t = Nachfrage[:,:Stunde]
k = Kategorien[:, :Kategorien]
l = ["DE","FR","NL"]

Kraftwerke

Wirkungsgrade = Dict(Kraftwerke[:,:Kategorie] .=> Kraftwerke[:,:Wirkungsgrad])
Wirkungsgrade
Brennstoffe = Dict(Kraftwerke[:,:Kategorie] .=> Kraftwerke[:,:Energieträger])
Brennstoffe

C = Kategorien[:, l]  #Kapazität 

function GK(i)

    f = Brennstoffe[i]
    η = Wirkungsgrade[i]
    p_f = Energieträger[1, f] #Preis je Brennstoff und Brennstoff hängt über f von Kraftwerkskategorie ab
    e_f = Energieträger[2, f] #Emissionsfaktor des Brennstoffes
    p_e = CO2_Preis[1, 1] #CO2-Preis
    #a_f = Verfügbarkeit[:, f] #Gebe die ganze Spalte je nach Brennstoffart aus

    p_el = (p_f / η) + (e_f / η) * p_e  #p_el = Grenzkosten

    return p_el
    push!(k, p_el)
end

costs=[]

for i in k
    GK(i)
    p_el = GK(i)
    push!(costs, p_el)
end
costs

Grenzkosten = Dict(Kraftwerke[:,:Kategorie] .=> costs[:,:])

