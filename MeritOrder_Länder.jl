#Extensions
using JuMP
using CPLEX
using XLSX, DataFrames

#Info: installierte Kapazität = maximal abrufbare Leistung; 
#      Leistung = eingesetzte Kapazität zu jeder Stunde

#Rufe die Excelliste "MeritOrder_Excel" und das entsprechende Tabellenblatt ab. Der Datentyp der Tabellenblatt-Inhalte wird ebenfalls definiert
Kapazität_df = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "Kapazität", infer_eltypes=true)...)
Kraftwerke_df = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "Kraftwerke", infer_eltypes=true)...)
Energieträger_df = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "Energieträger", infer_eltypes=true)...)
Nachfrage_df = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "Nachfrage")...) .|> float
CO2_Preis_df = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "CO2-Preis")...) .|> float
Wind_df = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "Wind", infer_eltypes=true)...)
Sonne_df = DataFrame(XLSX.readtable("MeritOrderLänder.xlsx", "Sonne", infer_eltypes=true)...)

# Größe der Dimensionen Zeit, Kraftwerke und Länder werden als Zahl bestimmt
t = size(Nachfrage_df,1)
k = size(Kraftwerke_df,1)
l = size(Nachfrage_df,2)

# Dimensionen Zeit, Kraftwerkskategorien und Länder werden als Sets/Vektoren ausgegeben
t_set = collect(1:t)
k_set = Kraftwerke_df[:,:Kategorie]
l_set = ["DE", "FR", "NL"]

# Dictionaries werden erstellt, welche benötigte Inhalte und Zuweisungen enthalten
Wirkungsgrade = Dict(Kraftwerke_df[:,:Kategorie] .=> Kraftwerke_df[:,:Wirkungsgrad])
Brennstoffe = Dict(Kraftwerke_df[:,:Kategorie] .=> Kraftwerke_df[:,:Energieträger])
Brennstoffkosten = Dict(Energieträger_df[:, :Energieträger] .=> Energieträger_df[:,:Brennstoffkosten])
Emissionsfaktor = Dict(Energieträger_df[:, :Energieträger] .=> Energieträger_df[:, :Emissionsfaktor])
availability = Dict(Kraftwerke_df[:, :Kategorie] .=> Kraftwerke_df[:, :Verfügbarkeit])
# Kapazitäten dicitionary wird je nach Land und Kraftwerkstyp erstellt
Kapazität = Dict()
    for p in l_set
        push!(Kapazität, p => Dict(Kraftwerke_df[:,:Kategorie] .=> Kapazität_df[:,p]),) 
    end
Kapazität

# Vorbereitung der Verfügbarkeit je Kraftwerkskategorie. Fossile Kraftwerke, Wasserkraft und Biomasse sind zu jeder Stunde zu 0,95% verfügbar. 
# Wind und Sonne sind in ihrer Verfügbarkeit abhängig von der Zeit im Jahr
wind(l_set) = Wind_df[:,l_set]
sonne(l_set) = Sonne_df[:,l_set]
fossil = fill(0.95, (t))
# Anlegen eines Dictionaries für die Verfügbarkeiten der Kraftwerke
Verfügbarkeit = Dict(
    "Kernenergie" => Dict(),
    "Braunkohle_0" => Dict(),
    "Braunkohle_-" => Dict(),
    "Braunkohle_+" => Dict(),
    "Steinkohle_0" => Dict(),
    "Steinkohle_-" => Dict(),
    "Steinkohle_+" => Dict(),
    "Erdgas_0" => Dict(),
    "Erdgas_-" => Dict(),
    "Erdgas_+" => Dict(),
    "Biomasse" => Dict(),
    "Wasserkraft" => Dict(),
    "Windkraft" => Dict(),
    "PV" => Dict()
)
    # Dicitionary Verfügbarkeit wird mit for Schleife gefüllt, je nach Kraftwerkskategorie 
    for c in k_set
        for p in l_set
            if availability[c] == "Fossil"
            push!(Verfügbarkeit[c], p => Dict(t_set .=> fossil))

            elseif availability[c] == "Wind"
            push!(Verfügbarkeit[c], p => Dict(t_set .=> wind(p)))
            
            else availability[c] == "Sonne"
            push!(Verfügbarkeit[c], p => Dict(t_set .=> sonne(p)))
            end
        end
    end        

Verfügbarkeit

# Mit Hilfe der Dictionaries werden die Grenzkosten der Kraftwerke berechnet
function GK(i)
    f = Brennstoffe[i] #Verwendeter Brennstoff je Kraftwerkskategorie
    η = Wirkungsgrade[i] #Wirkungsgrad je Kraftwerkskategorie
    p_f = Brennstoffkosten[f] #Preis je Brennstoff und Brennstoff hängt über f von Kraftwerkskategorie ab
    e_f = Emissionsfaktor[f] #Emissionsfaktor des Brennstoffes
    p_e = CO2_Preis_df[1, 1] #CO2-Preis
    #a_f = Verfügbarkeit[:, f] #Gebe die ganze Spalte je nach Brennstoffart aus

    p_el = (p_f / η) + (e_f / η) * p_e  #p_el = Grenzkosten

   return p_el
end

#Grenzkosten je Kraftwerkskategorie werden in eine Matrix "costs" eingefügt
costs=[]
    for i in Kraftwerke_df[:,:Kategorie]
        GK(i)
        p_el = GK(i)
        push!(costs, p_el)
    end
convert(Array{Float64, 1}, costs) # Datentyp für von Any zu Float geändert
costs

# Weiteres Dictionary wo die Kosten den verschiedenen Kraftwerken zugewiesen werden
Grenzkosten = Dict(k_set .=> costs)

#Zusammenfassung:
t
k
l
t_set
k_set
l_set
Wirkungsgrade #Nur um GK zu berechnen
Brennstoffe #Nur um GK zu berechnen
Brennstoffkosten #Nur um GK zu berechnen
Emissionsfaktor #Nur um GK zu berechnen

Grenzkosten #Brauchen wir im Modell

Nachfrage_df #Abhängig von Zeit und Land -> fürs Modell
Kapazität #Abhängig von Kategorie -> fürs Modell
Verfügbarkeit #Abhängig von Kategorie -> fürs Modell

#Zu optimierendes Modell wird erstellt
model = Model(CPLEX.Optimizer)
set_silent(model)

@variable(model, x[t in t_set, k in k_set , l in l_set] >= 0) # Abgerufene Leistung ist abhängig von der Zeit, dem Kraftwerk und des Landes  
@objective(model, Min, sum(Grenzkosten[k]*x[t,k,l] for t in t_set, k in k_set, l in l_set)) # Zielfunktion: Multipliziere für jede Kraftwerkskategorie die Grenzkosten mit der eingesetzten Leistung in jeder Stunde und abhängig vom Land -> Minimieren
@constraint(model, c1[t in t_set, l in l_set], sum(x[t,:,l]) == Nachfrage_df[t,l]) # Die Summe der Leistungen über die Kraftwerkskategorien je Stunde darf nicht größer sein als die Nachfrage 
@constraint(model, c2[t in t_set, k in k_set, l in l_set], x[t,k,l] .<= Kapazität[l][k]*Verfügbarkeit[k][l][t]) # Die Leistung je Kraftwerkskategorie muss kleiner sein als die Kapazität...
#...der Kraftwerkskategorie in dem betrachteten Land multipliziert mit der Verfügbarkeit -> Verwendung der Inhalte aus den Dictionaries

optimize!(model)

x_results = @show value.(x) #Matrix aller Leistungen. x_results hat drei Dimensionen
obj_value = @show objective_value(model) #Minimierte Gesamtkosten der Stromerzeugung im gesamten Jahr
el_price = @show shadow_price.(c1) #Strompreis in jeder Stunde des Jahres

# Die Ergebnisse werden als Excellisten exportiert, jedes Land als extra Excel. Inhalt: verwendete Leistung je Stunde und Kraftwerkskategorie
for z in l_set
    y_results = x_results[:,:,z]
    z_results = Array(y_results)
    Excelname = "Ergebnisse"*z*".xlsx"
    results = DataFrame(z_results, k_set)
    rm(Excelname, force=true) #Lösche die alte, bereits bestehende Excel-Ergebnisliste
    XLSX.writetable(Excelname, results) #Erstelle eine neue Ergebnisliste
end
