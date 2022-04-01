#Extensions
using JuMP
using CPLEX
using XLSX, DataFrames

#Info: installierte Kapazität = maximal abrufbare Leistung; 
#      Leistung = eingesetzte Kapazität zu jeder Stunde

#Rufe die Excelliste "MeritOrderLänder" und das entsprechende Tabellenblatt ab. Der Datentyp der Tabellenblatt-Inhalte wird ebenfalls definiert
Kapazität_df = DataFrame(XLSX.readtable("MeritOrderHandel.xlsx", "Kapazität", infer_eltypes=true)...)
Kraftwerke_df = DataFrame(XLSX.readtable("MeritOrderHandel.xlsx", "Kraftwerke", infer_eltypes=true)...)
Energieträger_df = DataFrame(XLSX.readtable("MeritOrderHandel.xlsx", "Energieträger", infer_eltypes=true)...)
Nachfrage_df = DataFrame(XLSX.readtable("MeritOrderHandel.xlsx", "Nachfrage")...) .|> float
CO2_Preis_df = DataFrame(XLSX.readtable("MeritOrderHandel.xlsx", "CO2-Preis")...) .|> float
Wind_df = DataFrame(XLSX.readtable("MeritOrderHandel.xlsx", "Wind", infer_eltypes=true)...)
Sonne_df = DataFrame(XLSX.readtable("MeritOrderHandel.xlsx", "Sonne", infer_eltypes=true)...)

# Größe der Dimensionen Zeit, Kraftwerke und Länder werden als Zahl bestimmt
t = size(Nachfrage_df,1)
k = size(Kraftwerke_df,1)
l = size(Nachfrage_df,2)

#Wenn weniger Stunden betrachtet werden sollen
t = 48

Nachfrage_df = Nachfrage_df[1:t,:]
Wind_df = Wind_df[1:t, :]
Sonne_df = Sonne_df[1:t, :]

# Dimensionen Zeit, Kraftwerkskategorien und Länder werden als Sets/Vektoren ausgegeben
t_set = collect(1:t)
k_set = Kraftwerke_df[:,:Kategorie]
l_set = String.(names(Nachfrage_df)) #Länderbezeichnungen als Vektor

# Dictionaries werden erstellt, welche benötigte Inhalte und Zuweisungen enthalten
Wirkungsgrade = Dict(k_set .=> Kraftwerke_df[:,:Wirkungsgrad])
Brennstoffe = Dict(k_set .=> Kraftwerke_df[:,:Energieträger])
Brennstoffkosten = Dict(Energieträger_df[:, :Energieträger] .=> Energieträger_df[:,:Brennstoffkosten])
Emissionsfaktor = Dict(Energieträger_df[:, :Energieträger] .=> Energieträger_df[:, :Emissionsfaktor])
availability = Dict(k_set .=> Kraftwerke_df[:, :Verfügbarkeit])
Effizienz = Dict(k_set .=> Kraftwerke_df[:, :Effizienz])

# Kapazitäten dicitionary wird je nach Land und Kraftwerkstyp erstellt
Kapazität = Dict()
    for p in l_set
        push!(Kapazität, p => Dict(k_set .=> Kapazität_df[:,p]),) 
    end
Kapazität
# Nachfrage dictionary wird je nach Land und Stunde erstellt
Nachfrage = Dict()
    for p in l_set
        push!(Nachfrage, p => Dict(t_set .=> Nachfrage_df[:,p]),)
    end
Nachfrage

# Vorbereitung der Verfügbarkeit je Kraftwerkskategorie. 
# Wind und Sonne sind in ihrer Verfügbarkeit abhängig von der Zeit im Jahr und vom Land
wind(l_set) = Wind_df[:,l_set]
sonne(l_set) = Sonne_df[:,l_set]

# Anlegen eines Dictionaries für die Verfügbarkeiten der Kraftwerke
Verfügbarkeit = Dict()
    for c in k_set
        push!(Verfügbarkeit, c => Dict(),)
    end
Verfügbarkeit

    # Dicitionary Verfügbarkeit wird mit for Schleife gefüllt, je nach Kraftwerkskategorie 
    for c in k_set
        for p in l_set
            if availability[c] == "Wind"
            push!(Verfügbarkeit[c], p => Dict(t_set .=> wind(p)))
            
            elseif availability[c] == "Sonne"
            push!(Verfügbarkeit[c], p => Dict(t_set .=> sonne(p)))

            else 
            push!(Verfügbarkeit[c], p => Dict(t_set .=> fill(availability[c],(t))))
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

    p_el = (p_f / η) + (e_f / η) * p_e  #p_el = Grenzkosten

   return p_el
end

#Grenzkosten je Kraftwerkskategorie werden in eine Matrix "costs" eingefügt
Grenzkosten = Dict()
    for i in k_set
        GK(i)
        p_el = GK(i)
        push!(Grenzkosten, i .=> p_el)
    end
Grenzkosten


#Zusammenfassung:
t_set
k_set
l_set

Grenzkosten #Brauchen wir im Modell
Nachfrage #Abhängig von Zeit und Land -> fürs Modell
Kapazität #Abhängig von Kategorie -> fürs Modell
Verfügbarkeit #Abhängig von Kategorie -> fürs Modell
Effizienz

#Zu optimierendes Modell wird erstellt
model = Model(CPLEX.Optimizer)
set_silent(model)

@variable(model, x[t in t_set, k in k_set , l in l_set] >= 0) # Abgerufene Leistung ist abhängig von der Zeit, dem Kraftwerk und des Landes  

@objective(model, Min, sum(Grenzkosten[k]*x[t,k,l] for t in t_set, k in k_set, l in l_set)) # Zielfunktion: Multipliziere für jede Kraftwerkskategorie die Grenzkosten mit der eingesetzten Leistung in jeder Stunde und abhängig vom Land -> Minimieren

@constraint(model, c1[t in t_set, l in l_set], sum(x[t,g,l] * Effizienz[g] for g in k_set) == Nachfrage[l][t] + sum(x[t,l,j] for j in l_set)) 

@constraint(model, c2[t in t_set, k in k_set, l in l_set], x[t,k,l] .<= Kapazität[l][k]*Verfügbarkeit[k][l][t]*Effizienz[k]) # Die Leistung je Kraftwerkskategorie muss kleiner sein als die Kapazität...
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
