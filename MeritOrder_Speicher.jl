#Extensions
using JuMP
using CPLEX
using XLSX, DataFrames

# Info: Installierte Kapazität = maximal abrufbare Leistung; 
#       Leistung = eingesetzte Kapazität zu jeder Stunde

# Rufe die Excelliste "MeritOrderSpeicher" und das entsprechende Tabellenblatt ab. Der Datentyp der Tabellenblattinhalte wird ebenfalls definiert
# MeritOrderSpeicher ist der Dateninput 
Kapazität_df = DataFrame(XLSX.readtable("MeritOrderSpeicher.xlsx", "Kapazität", infer_eltypes=true)...)
Kraftwerke_df = DataFrame(XLSX.readtable("MeritOrderSpeicher.xlsx", "Kraftwerke", infer_eltypes=true)...)
Energieträger_df = DataFrame(XLSX.readtable("MeritOrderSpeicher.xlsx", "Energieträger", infer_eltypes=true)...)
Nachfrage_df = DataFrame(XLSX.readtable("MeritOrderSpeicher.xlsx", "Nachfrage")...) .|> float
CO2_Preis_df = DataFrame(XLSX.readtable("MeritOrderSpeicher.xlsx", "CO2-Preis")...) .|> float
Wind_df = DataFrame(XLSX.readtable("MeritOrderSpeicher.xlsx", "Wind", infer_eltypes=true)...)
Sonne_df = DataFrame(XLSX.readtable("MeritOrderSpeicher.xlsx", "Sonne", infer_eltypes=true)...)
Effizienz_df = DataFrame(XLSX.readtable("MeritOrderSpeicher.xlsx", "Effizienz", infer_eltypes=true)...)

# Größe der Dimensionen Zeit, Kraftwerke und Länder werden als Zahl bestimmt
t = size(Nachfrage_df,1)
k = size(Kraftwerke_df,1)
l = size(Nachfrage_df,2)
n = size(Kapazität_df,2) 
s = n - l

# Wenn weniger Stunden betrachtet werden sollen hier eingeben, max. 8760
t = 35

# Die Tabellen Stromlast und Verfügbarkeit von Wind & Sonne wird auf den zu betrachtenden Zeitraum reduziert
Nachfrage_df = Nachfrage_df[1:t,:]
Wind_df = Wind_df[1:t, :]
Sonne_df = Sonne_df[1:t, :]

# Dimensionen Zeit, Kraftwerkskategorien und Länder werden als Sets/Vektoren ausgegeben
t_set = collect(1:t)
t_set_0 = collect(0:t)
k_set = Kraftwerke_df[:,:Kategorie]
l_set = String.(names(Nachfrage_df)) #Länderbezeichnungen als Vektor
s_set = ["Pumpspeicher"] #Speicher
n_set = String.(names(Kapazität_df)) #Länder und Speicher werden als Kraftwerke betrachtet -> Länder beim Handel und Speicher als Abnehmer oder Erzeuger


# Dictionaries werden erstellt, welche benötigte Inhalte und Zuweisungen enthalten
Wirkungsgrad = Dict(k_set .=> Kraftwerke_df[:,:Wirkungsgrad])
Brennstoffe = Dict(k_set .=> Kraftwerke_df[:,:Energieträger])
Brennstoffkosten = Dict(Energieträger_df[:, :Energieträger] .=> Energieträger_df[:,:Brennstoffkosten])
Emissionsfaktor = Dict(Energieträger_df[:, :Energieträger] .=> Energieträger_df[:, :Emissionsfaktor])
availability = Dict(k_set .=> Kraftwerke_df[:, :Verfügbarkeit])
Volumenfaktor = Dict(k_set .=> Kraftwerke_df[:, :Volumenfaktor])


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

Effizienz = Dict()
    for c in k_set
        if c in l_set
        push!(Effizienz, c => Dict(l_set .=> Effizienz_df[:,c]))
        
        else
        push!(Effizienz, c => Dict(l_set .=> 1))    
        end
    end

Effizienz


# Mit Hilfe der Dictionaries werden die Grenzkosten der Kraftwerke berechnet
function GK(i)
    f = Brennstoffe[i] #Verwendeter Brennstoff je Kraftwerkskategorie
    η = Wirkungsgrad[i] #Wirkungsgrad je Kraftwerkskategorie
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
s_set
n_set

Grenzkosten #Brauchen wir im Modell
Nachfrage #Abhängig von Zeit und Land -> fürs Modell
Kapazität #Abhängig von Kategorie -> fürs Modell
Verfügbarkeit #Abhängig von Kategorie -> fürs Modell
Effizienz

#Zu optimierendes Modell wird erstellt
model = direct_model(CPLEX.Optimizer())
set_silent(model)

@variable(model, x[t in t_set, k in k_set , n in n_set] >= 0) # Abgerufene Leistung ist abhängig von der Zeit, dem Kraftwerk und Land  
@variable(model, 0 <= y[t in t_set_0, s in s_set, l in l_set] <= Volumenfaktor[s] * Kapazität[l][s]) # Variable y überprüft das Speicherlevel: Darf nicht höher sein als installierte Kapazität * Volumenfaktor & muss größer Null sein
@objective(model, Min, sum(Grenzkosten[k]*x[t,k,n] for t in t_set, k in k_set, n in n_set)) # Zielfunktion: Multipliziere für jede Kraftwerkskategorie die Grenzkosten mit der eingesetzten Leistung in jeder Stunde und abhängig vom Land -> Minimieren
@constraint(model, c1[t in t_set, l in l_set], sum(x[t,k,l] * Effizienz[k][l] for k in k_set) == Nachfrage[l][t] + sum(x[t,l,j] for j in l_set) + sum(x[t,l,s] / Wirkungsgrad[s] for s in s_set)) # Summe der eingesetzten Leistung soll mit der Effizienz multipliziert werden (für eigenen Verbrauch ist die Effizienz 1, für Handel ist sie kleiner -> Grund Eigenverbrauch soll vorrangig passieren)...
# ... auf die eigene Nachfrage des Landes wird die Summe die exportiert wird draufgerechnet, da dies extra produziert wird. Das findet nur für Kraftwerke statt, die auch Länder sind. 
# ... Zusätzlich wird überschüssige Energie eines Landes eingespeichert. Die Einspeicherung wird mit einem Wirkungsgrad (Verluste) versehen und auf die Nachfrage addiert. 
# ... Die Ausspeicherung ist auf der linken Gleichheitszeichen im x enthalten, da die Ausspeicherung wie die Stromerzeugung eines Kraftwerkes behandelt wird.
@constraint(model, c2[t in t_set, k in k_set, l in l_set], x[t,k,l] .<= Kapazität[l][k]*Verfügbarkeit[k][l][t]) # Die Leistung je Kraftwerkskategorie muss kleiner sein als die Kapazität...
#...der Kraftwerkskategorie in dem betrachteten Land multipliziert mit der Verfügbarkeit -> Verwendung der Inhalte aus den Dictionaries. Speicher hier enthalten, diese werden im Falle der Ausspeicherung auf die zur Verfügung stehende Kapazität beschränkt
@constraint(model, c3[t in t_set, l in l_set, s in s_set], x[t,l,s] .<= Kapazität[l][s]) # Nebenbedingung 3 beschränkt die Einspeicherung auf die verfügbare Kapazität des Speichers je Land
@constraint(model, c4[t in t_set, s in s_set, l in l_set], y[t,s,l] == y[t-1,s,l] + x[t,l,s] - x[t,s,l]) # NB 4 gibt das Speicherlevel aus. Das Speicherlevel der betrachteten Stunde muss die Summe sein aus dem Level der vorherhigen Stunde + Einspeicherung - Ausspeicherung
@constraint(model, c5[s in s_set, l in l_set], y[0,s,l] == y[t,s,l]) # NB 5 sagt, dass das Speicherlevel zu Stunde Null gleich dem Speicherstand der letzten betrachteten Stunde sein muss
@constraint(model, c6[s in s_set, l in l_set], y[0,s,l] == 0.5*Volumenfaktor[s] * Kapazität[l][s]) # NB 5 sagt, dass das Speicherlevel zur Stunde Null der halben Kapazität entsprechen muss

optimize!(model)

x_results = @show value.(x) # Matrix aller abgerufenen Leistungen. x_results hat drei Dimensionen
y_results = @show value.(y) # Matrix aller abgerufenen Speicherstände
obj_value = @show objective_value(model) # Minimierte Gesamtkosten der Stromerzeugung im gesamten Jahr
el_price = @show shadow_price.(c1)*(-1) # Strompreis in jeder Stunde des Jahres


# Ausgabe der Ergebnisse je Land 
# Optimierter Kraftwerkseinsatz je Land mit Import (alles was aus anderen Ländern nach bspw. D kommt)
DE_df = DataFrame(Array(x_results[:,:,"DE"]), k_set) 
FR_df = DataFrame(Array(x_results[:,:,"FR"]), k_set)
NL_df = DataFrame(Array(x_results[:,:,"NL"]), k_set)
PL_df = DataFrame(Array(x_results[:,:,"PL"]), k_set)
SE_df = DataFrame(Array(x_results[:,:,"SE"]), k_set)
NO_df = DataFrame(Array(x_results[:,:,"NO"]), k_set)
# Einspeicherung aller Länder in deren Pumpspeicher wird angezeigt
Pumpspeicher = DataFrame(Array(x_results[:,:,"Pumpspeicher"]), k_set)
# Speicherstand am Anfang der betrachteten Stunde wird angezeigt je Land
Speicherstand = DataFrame(Array(y_results[:,"Pumpspeicher",:]), l_set) 

# Ausgabe der Strompreise je Land
Strompreise = DataFrame(Array(el_price[:,:]), l_set)

# Vorbereitung des Ergebnis-Outputs in Excel
# Dataframe für jedes Land wird erstellt, den Ergebnissen von oben werden die Exporte in die verschiedenen Länder angehängt, sowie die Einspeicherung und die Nachfrage des betrachteten Ladens  
DE = hcat(DE_df, DataFrame(
     hcat(FR_df[:,:DE], NL_df[:,:DE], PL_df[:,:DE], SE_df[:,:DE], NO_df[:,:DE], Pumpspeicher[:,:DE], Nachfrage_df[:,:DE], Strompreise[:,:DE]), 
     ["FR_ex", "NL_ex", "PL_ex", "SE_ex", "NO_ex", "Einspeicherung", "Nachfrage", "Strompreis"]))

FR = hcat(FR_df, DataFrame(
     hcat(DE_df[:,:FR], NL_df[:,:FR], PL_df[:,:FR], SE_df[:,:FR], NO_df[:,:FR], Pumpspeicher[:,:FR], Nachfrage_df[:,:FR], Strompreise[:,:FR]), 
     ["DE_ex", "NL_ex", "PL_ex", "SE_ex", "NO_ex", "Einspeicherung", "Nachfrage", "Strompreis"]))

NL = hcat(NL_df, DataFrame(
     hcat(DE_df[:,:NL], FR_df[:,:NL], PL_df[:,:NL], SE_df[:,:NL], NO_df[:,:NL], Pumpspeicher[:,:NL], Nachfrage_df[:,:NL], Strompreise[:,:NL]), 
     ["DE_ex", "FR_ex", "PL_ex", "SE_ex", "NO_ex", "Einspeicherung", "Nachfrage", "Strompreis"]))

PL = hcat(PL_df, DataFrame(
     hcat(DE_df[:,:PL], FR_df[:,:PL], NL_df[:,:PL], SE_df[:,:PL], NO_df[:,:PL], Pumpspeicher[:,:PL], Nachfrage_df[:,:PL], Strompreise[:,:PL]), 
     ["DE_ex", "FR_ex", "NL_ex", "SE_ex", "NO_ex", "Einspeicherung", "Nachfrage", "Strompreis"]))
   
SE = hcat(SE_df, DataFrame(
     hcat(DE_df[:,:SE], FR_df[:,:SE], NL_df[:,:SE], PL_df[:,:SE], NO_df[:,:SE], Pumpspeicher[:,:SE], Nachfrage_df[:,:SE], Strompreise[:,:SE]), 
     ["DE_ex", "FR_ex", "NL_ex", "PL_ex", "NO_ex", "Einspeicherung", "Nachfrage", "Strompreis"]))
   
NO = hcat(NO_df, DataFrame(
     hcat(DE_df[:,:NO], FR_df[:,:NO], NL_df[:,:NO], PL_df[:,:NO], SE_df[:,:NO], Pumpspeicher[:,:NO], Nachfrage_df[:,:NO], Strompreise[:,:NO]), 
     ["DE_ex", "FR_ex", "NL_ex", "PL_ex", "SE_ex", "Einspeicherung", "Nachfrage", "Strompreis"]))


# Export der vorbereiteten DataFrames in Excel
# Die Namen der Tabellenblätter müssen händisch erweitert werden, falls Länder & Speicher hinzugefügt werden
XLSX.writetable("Ergebnisse.xlsx", overwrite=true, 
        "DE" => DE,
        "FR" => FR, 
        "NL" => NL, 
        "PL" => PL,
        "SE" => SE,
        "NO" => NO,
        "Strompreise" => Strompreise, 
        "Pumpspeicher_Einspeicherung" => Pumpspeicher, 
        "Speicherstand" => Speicherstand,
        "Nachfrage" => Nachfrage_df)