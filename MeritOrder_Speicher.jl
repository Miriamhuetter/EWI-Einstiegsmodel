#Extensions
using JuMP
using CPLEX
using XLSX, DataFrames
using Plots, PlotlyJS


# Info: Installierte Kapazität = maximal abrufbare Leistung; 
#       Leistung = eingesetzte Kapazität zu jeder Stunde

# Rufe die Excelliste "MeritOrderSpeicher" und das entsprechende Tabellenblatt ab. Der Datentyp der Tabellenblattinhalte wird ebenfalls definiert
# MeritOrderSpeicher ist der Dateninput 
Kapazität_df = DataFrame(XLSX.readtable("MeritOrderSpeicher.xlsx", "Kapazität", infer_eltypes=true)...)
Kraftwerke_df = DataFrame(XLSX.readtable("MeritOrderSpeicher.xlsx", "Kraftwerke", infer_eltypes=true)...)
Volumenfaktor_df = DataFrame(XLSX.readtable("MeritOrderSpeicher.xlsx", "Volumenfaktor", infer_eltypes=true)...)
Effizienz_df = DataFrame(XLSX.readtable("MeritOrderSpeicher.xlsx", "Effizienz", infer_eltypes=true)...)
Energieträger_df = DataFrame(XLSX.readtable("MeritOrderSpeicher.xlsx", "Energieträger", infer_eltypes=true)...)
Nachfrage_df = DataFrame(XLSX.readtable("MeritOrderSpeicher.xlsx", "Nachfrage")...) .|> float
CO2_Preis = DataFrame(XLSX.readtable("MeritOrderSpeicher.xlsx", "CO2-Preis")...)[1,1] .|> float
Wind_df = DataFrame(XLSX.readtable("MeritOrderSpeicher.xlsx", "Wind", infer_eltypes=true)...)
Sonne_df = DataFrame(XLSX.readtable("MeritOrderSpeicher.xlsx", "Sonne", infer_eltypes=true)...)


# Größe der Dimensionen Zeit, Kraftwerke und Länder werden als Zahl bestimmt
t = size(Nachfrage_df,1)
k = size(Kraftwerke_df,1)
l = size(Nachfrage_df,2)
n = size(Kapazität_df,2) 
s = n - l

# Wenn weniger Stunden betrachtet werden sollen hier eingeben, max. 8760
t = 48

# Die Tabellen Stromlast und Verfügbarkeit von Wind & Sonne wird auf den zu betrachtenden Zeitraum reduziert
Nachfrage_df = Nachfrage_df[1:t,:]
Wind_df = Wind_df[1:t, :]
Sonne_df = Sonne_df[1:t, :]

# Dimensionen Zeit, Kraftwerkskategorien und Länder werden als Sets/Vektoren ausgegeben
t_set = collect(1:t)
k_set = Kraftwerke_df[:,:Kategorie]
l_set = String.(names(Nachfrage_df)) #Länderbezeichnungen als Vektor
s_set = ["Pumpspeicher", "Batteriespeicher", "Wasserstoffspeicher"] #Speicher
n_set = String.(names(Kapazität_df)) #Länder und Speicher werden als Kraftwerke betrachtet -> Länder beim Handel und Speicher als Abnehmer oder Erzeuger


# Dictionaries werden erstellt, welche benötigte Inhalte und Zuweisungen enthalten
Wirkungsgrad = Dict(k_set .=> Kraftwerke_df[:,:Wirkungsgrad])
Brennstoffe = Dict(k_set .=> Kraftwerke_df[:,:Energieträger])
Brennstoffkosten = Dict(Energieträger_df[:, :Energieträger] .=> Energieträger_df[:,:Brennstoffkosten])
Emissionsfaktor = Dict(Energieträger_df[:, :Energieträger] .=> Energieträger_df[:, :Emissionsfaktor])
availability = Dict(k_set .=> Kraftwerke_df[:, :Verfügbarkeit])

# Volumenfaktor dicitionary wird je nach Land und Speicher erstellt
Volumenfaktor = Dict()
    for s in s_set
        push!(Volumenfaktor, s => Dict(Volumenfaktor_df[:,:Land] .=> Volumenfaktor_df[:,s]),) 
    end
    push!(Volumenfaktor, "Speicherwasser" => Dict(Volumenfaktor_df[:,:Land] .=> Volumenfaktor_df[:,"Speicherwasser"])) 
Volumenfaktor

# Kapazitäten dicitionary wird je nach Land und Kraftwerkstyp erstellt
Kapazität = Dict()
    for p in n_set
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
        push!(Effizienz, c => Dict(Effizienz_df[:,:Exporteur] .=> Effizienz_df[:,c]))
        
        else
        push!(Effizienz, c => Dict(Effizienz_df[:,:Exporteur] .=> 1))    
        end
    end

Effizienz


# Mit Hilfe der Dictionaries werden die Grenzkosten der Kraftwerke berechnet
function GK(i)
    f = Brennstoffe[i] #Verwendeter Brennstoff je Kraftwerkskategorie
    η = Wirkungsgrad[i] #Wirkungsgrad je Kraftwerkskategorie
    p_f = Brennstoffkosten[f] #Preis je Brennstoff und Brennstoff hängt über f von Kraftwerkskategorie ab
    e_f = Emissionsfaktor[f] #Emissionsfaktor des Brennstoffes
    p_e = CO2_Preis #CO2-Preis

    p_el = (p_f / η) + (e_f / η) * p_e  #p_el = Grenzkosten
    e_el = (e_f / η)
   return p_el, e_el
end

#Grenzkosten je Kraftwerkskategorie werden in eine Dicitionary "Grenzkosten" reingepusht
Grenzkosten = Dict()
    for i in k_set
        p_el, e_el = GK(i)
        push!(Grenzkosten, i .=> p_el)
    end
Grenzkosten

#Emissionen je Kraftwerkskategorie werden in ein Dicitionary "Emissionsfaktor_elektisch" reingepusht -> Umrechnung von thermischen zu elektrischen Emissionen werden vorher in funktion GK(i) umgerechnet mittels Wirkungsgrad
Emissionsfaktor_el = Dict()
    for i in k_set
        p_el, e_el = GK(i)
        push!(Emissionsfaktor_el, i .=> e_el)
    end
Emissionsfaktor_el

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
spw = "Speicherwasser"

#Zu optimierendes Modell wird erstellt
model = direct_model(CPLEX.Optimizer())
set_silent(model)

@variable(model, x[t in t_set, k in k_set , n in n_set] >= 0) # Abgerufene Leistung ist abhängig von der Zeit, dem Kraftwerk und Land  
@variable(model, 0 <= y[t in t_set, s in s_set, l in l_set] <= Volumenfaktor[s][l] * Kapazität[l][s]) # Variable y überprüft das Speicherlevel: Darf nicht höher sein als installierte Kapazität * Volumenfaktor & muss größer Null sein
@variable(model, 0 <= sw[t in t_set, l in l_set] <= Volumenfaktor[spw][l] * Kapazität[l][spw])
@variable(model, z[t in t_set, l in l_set]) # Emissionen
@objective(model, Min, sum(Grenzkosten[k]*x[t,k,n] for t in t_set, k in k_set, n in n_set)) # Zielfunktion: Multipliziere für jede Kraftwerkskategorie die Grenzkosten mit der eingesetzten Leistung in jeder Stunde und abhängig vom Land -> Minimieren
@constraint(model, Bilanz[t in t_set, l in l_set], sum(x[t,k,l] * Effizienz[k][l] for k in k_set) == Nachfrage[l][t] + sum(x[t,l,j] for j in l_set) + sum(x[t,l,s] / Wirkungsgrad[s] for s in s_set)) # Summe der eingesetzten Leistung soll mit der Effizienz multipliziert werden (für eigenen Verbrauch ist die Effizienz 1, für Handel ist sie kleiner -> Grund Eigenverbrauch soll vorrangig passieren)...
# ... auf die eigene Nachfrage des Landes wird die Summe die exportiert wird draufgerechnet, da dies extra produziert wird. Das findet nur für Kraftwerke statt, die auch Länder sind. 
# ... Zusätzlich wird überschüssige Energie eines Landes eingespeichert. Die Einspeicherung wird mit einem Wirkungsgrad (Verluste) versehen und auf die Nachfrage addiert. 
# ... Die Ausspeicherung ist auf der linken Gleichheitszeichen im x enthalten, da die Ausspeicherung wie die Stromerzeugung eines Kraftwerkes behandelt wird.
@constraint(model, Kapazität_Kraftwerke[t in t_set, k in k_set, l in l_set], x[t,k,l] .<= Kapazität[l][k]*Verfügbarkeit[k][l][t]) # Die Leistung je Kraftwerkskategorie muss kleiner sein als die Kapazität...
#...der Kraftwerkskategorie in dem betrachteten Land multipliziert mit der Verfügbarkeit -> Verwendung der Inhalte aus den Dictionaries. Speicher hier enthalten, diese werden im Falle der Ausspeicherung auf die zur Verfügung stehende Kapazität beschränkt
@constraint(model, Kapazität_Speicher[t in t_set, l in l_set, s in s_set], x[t,l,s] .<= Kapazität[s][l]) # Nebenbedingung 3 beschränkt die Einspeicherung auf die verfügbare Kapazität des Speichers je Land
@constraint(model, Speicherstand_1[t in t_set[2:end], s in s_set, l in l_set], y[t,s,l] == y[t-1,s,l] + x[t-1,l,s] - x[t-1,s,l]) # NB 4 gibt das Speicherlevel aus. Das Speicherlevel der betrachteten Stunde muss die Summe sein aus dem Level der vorherhigen Stunde + Einspeicherung - Ausspeicherung
@constraint(model, Speicherstand_2[s in s_set, l in l_set], y[1,s,l] == y[t,s,l] + x[t,l,s] - x[t,s,l]) # NB 5 sagt, dass das Speicherlevel zu Stunde 1 gleich dem Speicherstand der letzten betrachteten Stunde sein muss
@constraint(model, Speicherstand_3[s in s_set, l in l_set], y[1,s,l] == 0.5*Volumenfaktor[s][l] * Kapazität[l][s]) # NB 5 sagt, dass das Speicherlevel zur Stunde Null der halben Kapazität entsprechen muss
@constraint(model, Emissionen[t in t_set, l in l_set], z[t,l] == sum(x[t,k,l] * Emissionsfaktor_el[k] for k in k_set))

@constraint(model, Speicherstand_1W[t in t_set[2:end], l in l_set], sw[t,l] == sw[t-1,l] - x[t-1,spw,l] + 0.0006*Volumenfaktor[spw][l] * Kapazität[l][spw]) # NB 4 gibt das Speicherlevel aus. Das Speicherlevel der betrachteten Stunde muss die Summe sein aus dem Level der vorherhigen Stunde + Einspeicherung - Ausspeicherung
@constraint(model, Speicherstand_2W[l in l_set], sw[1,l] == sw[t,l] - x[t,spw,l] + 0.0006*Volumenfaktor[spw][l] * Kapazität[l][spw]) # NB 5 sagt, dass das Speicherlevel zu Stunde 1 gleich dem Speicherstand der letzten betrachteten Stunde sein muss
@constraint(model, Speicherstand_3W[l in l_set], sw[1,l] == 0.5*Volumenfaktor[spw][l] * Kapazität[l][spw]) # NB 5 sagt, dass das Speicherlevel zur Stunde Null der halben Kapazität entsprechen muss

optimize!(model)
termination_status(model)

x_results = @show value.(x) # Matrix aller abgerufenen Leistungen. x_results hat drei Dimensionen
y_results = @show value.(y) # Matrix aller abgerufenen Speicherstände
z_results = @show value.(z)
sw_results = @show value.(sw)
obj_value = @show objective_value(model) # Minimierte Gesamtkosten der Stromerzeugung im gesamten Jahr
el_price = @show shadow_price.(Bilanz)*(-1) # Strompreis in jeder Stunde des Jahres

Ueberschriften = ["Kernenergie", "Braunkohle_+", "Braunkohle_0", "Braunkohle_-", "Steinkohle_+", "Steinkohle_0", "Steinkohle_-", "Erdgas_+", "Erdgas_0", "Erdgas_-", "Erdöl", "Windkraft", "PV", "Biomasse", "Laufwasser", "Speicherwasser", "DE_im", "FR_im", "NL_im", "PL_im", "SE_im", "NO_im", "AT_im", "Pumpspeicher_Ausspeicherung", "Batteriespeicher_Ausspeicherung", "Wasserstoffspeicher_Ausspeicherung"]

# Ausgabe der Ergebnisse je Land 
# Optimierter Kraftwerkseinsatz je Land mit Import (alles was aus anderen Ländern nach bspw. D kommt)
DE_df = DataFrame(Array(x_results[:,:,"DE"]), Ueberschriften) 
FR_df = DataFrame(Array(x_results[:,:,"FR"]), Ueberschriften)
NL_df = DataFrame(Array(x_results[:,:,"NL"]), Ueberschriften)
PL_df = DataFrame(Array(x_results[:,:,"PL"]), Ueberschriften)
SE_df = DataFrame(Array(x_results[:,:,"SE"]), Ueberschriften)
NO_df = DataFrame(Array(x_results[:,:,"NO"]), Ueberschriften)
AT_df = DataFrame(Array(x_results[:,:,"AT"]), Ueberschriften)
# Einspeicherung aller Länder in deren Pumpspeicher wird angezeigt
Pumpspeicher = DataFrame(Array(x_results[:,:,"Pumpspeicher"]), k_set)
# Einspeicherung aller Länder in deren Batteriespeicher wird angezeigt
Batteriespeicher = DataFrame(Array(x_results[:,:,"Batteriespeicher"]), k_set)
# Einspeicherung aller Länder in deren Wasserstoffspeicher wird angezeigt
Wasserstoffspeicher = DataFrame(Array(x_results[:,:,"Wasserstoffspeicher"]), k_set)

# Speicherstand am Anfang der betrachteten Stunde wird angezeigt je Land
PS_Speicherstand = DataFrame(Array(y_results[:,"Pumpspeicher",:]), l_set) 

BS_Speicherstand = DataFrame(Array(y_results[:,"Batteriespeicher",:]), l_set)

WS_Speicherstand = DataFrame(Array(y_results[:,"Wasserstoffspeicher",:]), l_set)

SW_Speicherstand = DataFrame(Array(sw_results[:,:]), l_set)

# Ausgabe der Strompreise je Land
Strompreise = DataFrame(Array(el_price[:,:]), l_set)

# Ausgabe der Emissionen je Stunde und Land
Emissionen = DataFrame(Array(z_results[:,:]), l_set)

# Vorbereitung des Ergebnis-Outputs in Excel
# Dataframe für jedes Land wird erstellt, den Ergebnissen von oben werden die Exporte in die verschiedenen Länder angehängt, sowie die Einspeicherung und die Nachfrage des betrachteten Ladens  
DE = hcat(DE_df, DataFrame(
     hcat(FR_df[:,:DE_im], NL_df[:,:DE_im], PL_df[:,:DE_im], SE_df[:,:DE_im], NO_df[:,:DE_im], AT_df[:,:DE_im], Pumpspeicher[:,:DE], PS_Speicherstand[:,:DE], Batteriespeicher[:,:DE], BS_Speicherstand[:,:DE], Wasserstoffspeicher[:,:DE], WS_Speicherstand[:,:DE], Nachfrage_df[:,:DE], Strompreise[:,:DE], Emissionen[:,:DE]), 
     ["FR_ex", "NL_ex", "PL_ex", "SE_ex", "NO_ex", "AT_ex", "PS_Einspeicherung", "PS_Speicherstand", "BS_Einspeicherung", "BS_Speicherstand", "WS_Einspeicherung", "WS_Speicherstand", "Nachfrage", "Strompreis", "Emissionen"]))

FR = hcat(FR_df, DataFrame(
     hcat(DE_df[:,:FR_im], NL_df[:,:FR_im], PL_df[:,:FR_im], SE_df[:,:FR_im], NO_df[:,:FR_im], AT_df[:,:FR_im], Pumpspeicher[:,:FR], PS_Speicherstand[:,:FR], Batteriespeicher[:,:FR], BS_Speicherstand[:,:FR], Wasserstoffspeicher[:,:FR], WS_Speicherstand[:,:FR], Nachfrage_df[:,:FR], Strompreise[:,:FR], Emissionen[:,:FR]), 
     ["DE_ex", "NL_ex", "PL_ex", "SE_ex", "NO_ex", "AT_ex", "PS_Einspeicherung", "PS_Speicherstand", "BS_Einspeicherung", "BS_Speicherstand", "WS_Einspeicherung", "WS_Speicherstand", "Nachfrage", "Strompreis", "Emissionen"]))

NL = hcat(NL_df, DataFrame(
     hcat(DE_df[:,:NL_im], FR_df[:,:NL_im], PL_df[:,:NL_im], SE_df[:,:NL_im], NO_df[:,:NL_im],  AT_df[:,:NL_im], Pumpspeicher[:,:NL], PS_Speicherstand[:,:NL], Batteriespeicher[:,:NL], BS_Speicherstand[:,:NL], Wasserstoffspeicher[:,:NL], WS_Speicherstand[:,:NL], Nachfrage_df[:,:NL], Strompreise[:,:NL], Emissionen[:,:NL]), 
     ["DE_ex", "FR_ex", "PL_ex", "SE_ex", "NO_ex", "AT_ex", "PS_Einspeicherung", "PS_Speicherstand", "BS_Einspeicherung", "BS_Speicherstand", "WS_Einspeicherung", "WS_Speicherstand", "Nachfrage", "Strompreis", "Emissionen"]))

PL = hcat(PL_df, DataFrame(
     hcat(DE_df[:,:PL_im], FR_df[:,:PL_im], NL_df[:,:PL_im], SE_df[:,:PL_im], NO_df[:,:PL_im],  AT_df[:,:PL_im], Pumpspeicher[:,:PL], PS_Speicherstand[:,:PL], Batteriespeicher[:,:PL], BS_Speicherstand[:,:PL], Wasserstoffspeicher[:,:PL], WS_Speicherstand[:,:PL], Nachfrage_df[:,:PL], Strompreise[:,:PL], Emissionen[:,:PL]), 
     ["DE_ex", "FR_ex", "NL_ex", "SE_ex", "NO_ex", "AT_ex", "PS_Einspeicherung", "PS_Speicherstand", "BS_Einspeicherung", "BS_Speicherstand", "WS_Einspeicherung", "WS_Speicherstand", "Nachfrage", "Strompreis", "Emissionen"]))
   
SE = hcat(SE_df, DataFrame(
     hcat(DE_df[:,:SE_im], FR_df[:,:SE_im], NL_df[:,:SE_im], PL_df[:,:SE_im], NO_df[:,:SE_im],  AT_df[:,:SE_im], Pumpspeicher[:,:SE], PS_Speicherstand[:,:SE], Batteriespeicher[:,:SE], BS_Speicherstand[:,:SE], Wasserstoffspeicher[:,:SE], WS_Speicherstand[:,:SE], Nachfrage_df[:,:SE], Strompreise[:,:SE], Emissionen[:,:SE]), 
     ["DE_ex", "FR_ex", "NL_ex", "PL_ex", "NO_ex", "AT_ex", "PS_Einspeicherung", "PS_Speicherstand", "BS_Einspeicherung", "BS_Speicherstand", "WS_Einspeicherung", "WS_Speicherstand", "Nachfrage", "Strompreis", "Emissionen"]))
   
NO = hcat(NO_df, DataFrame(
     hcat(DE_df[:,:NO_im], FR_df[:,:NO_im], NL_df[:,:NO_im], PL_df[:,:NO_im], SE_df[:,:NO_im],  AT_df[:,:NO_im], Pumpspeicher[:,:NO], PS_Speicherstand[:,:NO], Batteriespeicher[:,:NO], BS_Speicherstand[:,:NO], Wasserstoffspeicher[:,:NO], WS_Speicherstand[:,:NO], Nachfrage_df[:,:NO], Strompreise[:,:NO], Emissionen[:,:NO]), 
     ["DE_ex", "FR_ex", "NL_ex", "PL_ex", "SE_ex", "AT_ex", "PS_Einspeicherung", "PS_Speicherstand", "BS_Einspeicherung", "BS_Speicherstand", "WS_Einspeicherung", "WS_Speicherstand", "Nachfrage", "Strompreis", "Emissionen"]))

AT = hcat(AT_df, DataFrame(
     hcat(DE_df[:,:AT_im], FR_df[:,:AT_im], NL_df[:,:AT_im], PL_df[:,:AT_im], SE_df[:,:AT_im],  NO_df[:,:AT_im], Pumpspeicher[:,:AT], PS_Speicherstand[:,:AT], Batteriespeicher[:,:AT], BS_Speicherstand[:,:AT], Wasserstoffspeicher[:,:AT], WS_Speicherstand[:,:AT], Nachfrage_df[:,:AT], Strompreise[:,:AT], Emissionen[:,:AT]), 
     ["DE_ex", "FR_ex", "NL_ex", "PL_ex", "SE_ex", "NO_ex", "PS_Einspeicherung", "PS_Speicherstand", "BS_Einspeicherung", "BS_Speicherstand", "WS_Einspeicherung", "WS_Speicherstand", "Nachfrage", "Strompreis", "Emissionen"]))


# Export der vorbereiteten DataFrames in Excel
# Die Namen der Tabellenblätter müssen händisch erweitert werden, falls Länder & Speicher hinzugefügt werden
XLSX.writetable("Ergebnisse.xlsx", overwrite=true, 
        "DE" => DE,
        "FR" => FR, 
        "NL" => NL, 
        "PL" => PL,
        "SE" => SE,
        "NO" => NO,
        "AT" => AT,
        "Strompreise" => Strompreise, 
        "Emissionen" => Emissionen,
        "Nachfrage" => Nachfrage_df,
        "Speicherwasser" => SW_Speicherstand
        #"PS_Einspeicherung" => Pumpspeicher, 
        #"PS_Speicherstand" => PS_Speicherstand,
        #"BS_Einspeicherung" => Batteriespeicher,
        #"BS_Speicherstand" => BS_Speicherstand,
        #"WS_Einspeicherung" => Wasserstoffspeicher,
        #"WS_Speicherstand" => WS_Speicherstand,
)