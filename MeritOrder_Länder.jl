#Extensions
using JuMP
using CPLEX
using XLSX, DataFrames

#Rufe die Excelliste "MeritOrder_Excel" und das entsprechende Tabellenblatt ab. Der Datentyp der Tabellenblatt-Inhalte wird ebenfalls definiert
Kapazität = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "Kapazität", infer_eltypes=true)...)
Kraftwerke = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "Kraftwerke", infer_eltypes=true)...)
Energieträger = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "Energieträger", infer_eltypes=true)...)
Nachfrage = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "Nachfrage")...) .|> float
CO2_Preis = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "CO2-Preis")...) .|> float
Verfügbarkeit = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "Verfügbarkeit", infer_eltypes=true)...)

k_set = Kraftwerke[:,:Kategorie]


t = size(Nachfrage,1)
k = size(Kraftwerke,1)
l = size(Nachfrage,2)


Wirkungsgrade = Dict(Kraftwerke[:,:Kategorie] .=> Kraftwerke[:,:Wirkungsgrad])
Wirkungsgrade
Brennstoffe = Dict(Kraftwerke[:,:Kategorie] .=> Kraftwerke[:,:Energieträger])
Brennstoffe 

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

for i in Kraftwerke[:,:Kategorie]
    GK(i)
    p_el = GK(i)
    push!(costs, p_el)
end
convert(Array{Float64, 1}, costs)
costs
headers = ["Kategorien","Grenzkosten"]
Grenzkosten = DataFrame(hcat(Kraftwerke[:,:Kategorie], costs[:,:]), headers)


#Zusammenfassung:
t
k
l
Kraftwerke
Kapazität
Grenzkosten
Nachfrage


#Verfügbarkeit =

model = Model(CPLEX.Optimizer)
set_silent(model)

@variable(model, x[1:t,1:k,1:l] >= 0)
@objective(model, Min, sum(costs[k]*x[t,k,l] for t in 1:t, k in 1:k, l in 1:l))
@constraint(model, c1[t=1:t,l=1:l], sum(x[t,:,l]) == Nachfrage[t,l])
@constraint(model, c2[t=1:t,k=1:k,l=1:l], x[t,k,l] .<= Kapazität[k, l])

optimize!(model)

x_results = @show value.(x) #Matrix aller Leistungen
obj_value = @show objective_value(model) #Minimierte Gesamtkosten der Stromerzeugung im gesamten Jahr
el_price = @show shadow_price.(c1) #Strompreis in jeder Stunde des Jahres
