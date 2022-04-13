#Extensions
using JuMP
using CPLEX
using XLSX, DataFrames


# Info: Installierte Kapazität = maximal abrufbare Leistung; 
#       Leistung = eingesetzte Kapazität zu jeder Stunde

# Rufe die Excelliste "MeritOrderSpeicher" und das entsprechende Tabellenblatt ab. Der Datentyp der Tabellenblattinhalte wird ebenfalls definiert
# MeritOrderSpeicher ist der Dateninput 
Kapazität_df = DataFrame(XLSX.readtable("MeritOrderJahre.xlsx", "Kapazität", infer_eltypes=true)...)
Kraftwerke_df = DataFrame(XLSX.readtable("MeritOrderJahre.xlsx", "Kraftwerke", infer_eltypes=true)...)
Volumenfaktor_df = DataFrame(XLSX.readtable("MeritOrderJahre.xlsx", "Volumenfaktor", infer_eltypes=true)...)
Effizienz_df = DataFrame(XLSX.readtable("MeritOrderJahre.xlsx", "Effizienz", infer_eltypes=true)...)
Energieträger_df = DataFrame(XLSX.readtable("MeritOrderJahre.xlsx", "Energieträger", infer_eltypes=true)...)
Nachfrage_df = DataFrame(XLSX.readtable("MeritOrderJahre.xlsx", "Nachfrage")...) .|> float
CO2_Preis_df = DataFrame(XLSX.readtable("MeritOrderJahre.xlsx", "CO2-Preis")...) .|> float
Wind_df = DataFrame(XLSX.readtable("MeritOrderJahre.xlsx", "Wind", infer_eltypes=true)...)
Sonne_df = DataFrame(XLSX.readtable("MeritOrderJahre.xlsx", "Sonne", infer_eltypes=true)...)


# Größe der Dimensionen Zeit, Kraftwerke und Länder werden als Zahl bestimmt
#t = size(Nachfrage_df,1)
k = size(Kraftwerke_df,1)
l = size(Nachfrage_df,2)-2
n = size(Kapazität_df,2)-2 
s = n - l # Anzahl Speicher
spw = "Speicherwasser"

# Wenn weniger Stunden betrachtet werden sollen hier eingeben, max. 8760
t = 12

# Die Tabellen Stromlast und Verfügbarkeit von Wind & Sonne wird auf den zu betrachtenden Zeitraum reduziert
#Nachfrage_df = Nachfrage_df[1:t,:]
#Wind_df = Wind_df[1:t, :]
#Sonne_df = Sonne_df[1:t, :]

# Dimensionen Zeit, Kraftwerkskategorien und Länder werden als Sets/Vektoren ausgegeben
t_set = collect(1:t)
k_set = Kraftwerke_df[:,:Kategorie]
l_set = String.(names(Nachfrage_df))[3:end] #Länderbezeichnungen als Vektor
s_set = ["Pumpspeicher", "Batteriespeicher", "Wasserstoffspeicher"] #Speicher
n_set = String.(names(Kapazität_df))[3:end] #Länder und Speicher werden als Kraftwerke betrachtet -> Länder beim Handel und Speicher als Abnehmer oder Erzeuger
j_set = CO2_Preis_df[:,:Jahr] .|> Int64

# Dictionaries werden erstellt, welche benötigte Inhalte und Zuweisungen enthalten
## Dictionaries die unanbängig vom Jahr sind
Wirkungsgrad = Dict(k_set .=> Kraftwerke_df[:,:Wirkungsgrad])
Brennstoffe = Dict(k_set .=> Kraftwerke_df[:,:Energieträger])
availability = Dict(k_set .=> Kraftwerke_df[:, :Verfügbarkeit])

# Dictionaries die abhängig vom Jahr sind
CO2Preis = Dict(CO2_Preis_df[:, :Jahr] .=> CO2_Preis_df[:, :Preis])

Emissionsfaktor = Dict()
    for j in j_set
        ef = filter(row -> row.Jahr == j, Energieträger_df) # Es wird zuerst das Jahr gefiltert und die Brennstoffkosten je Jahr und Energieträger dann in das Dictionary Brennstoffkosten gepusht
        push!(Emissionsfaktor, j => Dict(ef[:, :Energieträger] .=> ef[:,:Emissionsfaktor]))
    end
Emissionsfaktor


Brennstoffkosten = Dict()
    for j in j_set
        bf = filter(row -> row.Jahr == j, Energieträger_df) # Es wird zuerst das Jahr gefiltert und die Brennstoffkosten je Jahr und Energieträger dann in das Dictionary Brennstoffkosten gepusht
        push!(Brennstoffkosten, j => Dict(bf[:, :Energieträger] .=> bf[:,:Brennstoffkosten]))
    end
Brennstoffkosten


# Volumenfaktor dicitionary wird je nach Jahr, Land und Speicher erstellt
Volumenfaktor = Dict()
    for j in j_set
        push!(Volumenfaktor, j => Dict(),)
    end
Volumenfaktor # Leeres Dictionary, das erstmal nach Jahren aufgegliedert ist
    
    for j in j_set
        vf = filter(row -> row.Jahr == j, Volumenfaktor_df) # Es wird nach dem betrachteten Jahr gefiltert
        for s in s_set # Parameter s geht alle Speicherarten durch
            push!(Volumenfaktor[j], s => Dict(vf[:,:Land] .=> vf[:,s]),) # jedem Land (im Input in den Zeilen stehend) wird der Volumenfaktor des Speichers (Spalte) zugewiesen. Das wird danach in das Dictionary mit den zuvor erstellten Jahren gepusht
        end
        push!(Volumenfaktor[j], "Speicherwasser" => Dict(vf[:,:Land] .=> vf[:,"Speicherwasser"])) 
    end
Volumenfaktor

# Kapazitäten dicitionary wird je nach Land und Kraftwerkstyp erstellt
Kapazität = Dict()
    for j in j_set
        push!(Kapazität, j => Dict(),)
    end
Kapazität

    for j in j_set
        kf = filter(row -> row.Jahr == j, Kapazität_df)
        for n in n_set # n in Nachfrager (Spalten)
            push!(Kapazität[j], n => Dict(kf[:, :Technologien] .=> kf[:,n]),) # Die Technologien sind die Anbieter
    end
end
Kapazität
#Kapazität[2020]["DE"]["Kernenergie"]

Effizienz = Dict()
    for j in j_set
        push!(Effizienz, j => Dict(),)
    end
Effizienz

for j in j_set
    ez = filter(row -> row.Jahr == j, Effizienz_df)
    for c in k_set
        if c in l_set
        push!(Effizienz[j], c => Dict(ez[:,:Exporteur] .=> ez[:,c]))
        
        else
            push!(Effizienz[j], c => Dict(ez[:,:Exporteur] .=> 1))    
        end
    end
end
Effizienz

# Nachfrage dictionary wird je nach Land und Stunde erstellt
Nachfrage = Dict()
    for j in j_set
        push!(Nachfrage, j => Dict(),)
    end
Nachfrage

    for j in j_set
        nf = filter(row -> row.Jahr == j, Nachfrage_df)
    for l in l_set
        push!(Nachfrage[j], l => Dict(nf[:,:Stunde] .=> nf[:,l]),)
    end
end
Nachfrage

# Vorbereitung der Verfügbarkeit je Kraftwerkskategorie. 
# Wind und Sonne sind in ihrer Verfügbarkeit abhängig von der Zeit im Jahr und vom Land
# Anlegen eines Dictionaries für die Verfügbarkeiten der Kraftwerke
Verfügbarkeit = Dict()
    for j in j_set
        push!(Verfügbarkeit, j => Dict(),)
        for c in k_set
            push!(Verfügbarkeit[j], c => Dict(),)
        end
    end
Verfügbarkeit

    # Dicitionary Verfügbarkeit wird mit for Schleife gefüllt, je nach Kraftwerkskategorie 
    for j in j_set
        wv = filter(row -> row.Jahr == j, Wind_df)
        sv = filter(row -> row.Jahr == j, Sonne_df)
    for c in k_set
        for l in l_set
            if availability[c] == "Wind"
            push!(Verfügbarkeit[j][c], l => Dict(wv[:,:Stunde] .=> wv[:,l]))
            
            elseif availability[c] == "Sonne"
            push!(Verfügbarkeit[j][c], l => Dict(sv[:,:Stunde] .=> sv[:,l]))

            else 
            push!(Verfügbarkeit[j][c], l => Dict(t_set .=> fill(availability[c],(t))))
            end
        end
    end    
end    
Verfügbarkeit
#Verfügbarkeit[2020]["Windkraft"]["FR"][3] -> Test

# Mit Hilfe der Dictionaries werden die Grenzkosten der Kraftwerke berechnet
function GK(j, i)
    f = Brennstoffe[i] #Verwendeter Brennstoff je Kraftwerkskategorie
    η = Wirkungsgrad[i] #Wirkungsgrad je Kraftwerkskategorie
    p_f = Brennstoffkosten[j][f] #Preis je Brennstoff und Brennstoff hängt über f von Kraftwerkskategorie ab
    e_f = Emissionsfaktor[j][f] #Emissionsfaktor des Brennstoffes
    p_e = CO2Preis[j] #CO2-Preis

    p_el = (p_f / η) + (e_f / η) * p_e  #p_el = Grenzkosten
    e_el = (e_f / η)
   return p_el, e_el
end

#Grenzkosten je Kraftwerkskategorie werden in eine Dicitionary "Grenzkosten" reingepusht
Grenzkosten = Dict()
    for j in j_set
        push!(Grenzkosten, j => Dict(),)
    end
Grenzkosten

    for j in j_set
        for i in k_set
            p_el, e_el = GK(j, i)
            push!(Grenzkosten[j], i .=> p_el)
        end
    end
Grenzkosten

#Emissionen je Kraftwerkskategorie werden in ein Dicitionary "Emissionsfaktor_elektisch" reingepusht -> Umrechnung von thermischen zu elektrischen Emissionen werden vorher in funktion GK(i) umgerechnet mittels Wirkungsgrad
Emissionsfaktor_el = Dict()
    for j in j_set
        push!(Emissionsfaktor_el, j => Dict(),)
    end
Emissionsfaktor_el

    for j in j_set
        for i in k_set
            p_el, e_el = GK(j, i)
            push!(Emissionsfaktor_el[j], i .=> e_el)
        end
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

for j in j_set
#Zu optimierendes Modell wird erstellt
model = direct_model(CPLEX.Optimizer())
set_silent(model)

@variable(model, x[t in t_set, k in k_set , n in n_set] >= 0) # Abgerufene Leistung ist abhängig von der Zeit, dem Kraftwerk und Land  
@variable(model, 0 <= y[t in t_set, s in s_set, l in l_set] <= Volumenfaktor[j][s][l] * Kapazität[j][l][s]) # Variable y überprüft das Speicherlevel: Darf nicht höher sein als installierte Kapazität * Volumenfaktor & muss größer Null sein
@variable(model, 0 <= sw[t in t_set, l in l_set] <= Volumenfaktor[j][spw][l] * Kapazität[j][l][spw])
@variable(model, z[t in t_set, l in l_set]) # Emissionen
@objective(model, Min, sum(Grenzkosten[j][k]*x[t,k,n] for t in t_set, k in k_set, n in n_set)) # Zielfunktion: Multipliziere für jede Kraftwerkskategorie die Grenzkosten mit der eingesetzten Leistung in jeder Stunde und abhängig vom Land -> Minimieren
@constraint(model, Bilanz[t in t_set, l in l_set], sum(x[t,k,l] * Effizienz[j][k][l] for k in k_set) == Nachfrage[j][l][t] + sum(x[t,l,j] for j in l_set) + sum(x[t,l,s] / Wirkungsgrad[s] for s in s_set)) # Summe der eingesetzten Leistung soll mit der Effizienz multipliziert werden (für eigenen Verbrauch ist die Effizienz 1, für Handel ist sie kleiner -> Grund Eigenverbrauch soll vorrangig passieren)...
# ... auf die eigene Nachfrage des Landes wird die Summe die exportiert wird draufgerechnet, da dies extra produziert wird. Das findet nur für Kraftwerke statt, die auch Länder sind. 
# ... Zusätzlich wird überschüssige Energie eines Landes eingespeichert. Die Einspeicherung wird mit einem Wirkungsgrad (Verluste) versehen und auf die Nachfrage addiert. 
# ... Die Ausspeicherung ist auf der linken Gleichheitszeichen im x enthalten, da die Ausspeicherung wie die Stromerzeugung eines Kraftwerkes behandelt wird.
@constraint(model, Kapazität_Kraftwerke[t in t_set, k in k_set, l in l_set], x[t,k,l] .<= Kapazität[j][l][k]*Verfügbarkeit[j][k][l][t]) # Die Leistung je Kraftwerkskategorie muss kleiner sein als die Kapazität...
#...der Kraftwerkskategorie in dem betrachteten Land multipliziert mit der Verfügbarkeit -> Verwendung der Inhalte aus den Dictionaries. Speicher hier enthalten, diese werden im Falle der Ausspeicherung auf die zur Verfügung stehende Kapazität beschränkt
@constraint(model, Kapazität_Speicher[t in t_set, l in l_set, s in s_set], x[t,l,s] .<= Kapazität[j][s][l]) # Nebenbedingung 3 beschränkt die Einspeicherung auf die verfügbare Kapazität des Speichers je Land
@constraint(model, Speicherstand_1[t in t_set[2:end], s in s_set, l in l_set], y[t,s,l] == y[t-1,s,l] + x[t-1,l,s] - x[t-1,s,l]) # NB 4 gibt das Speicherlevel aus. Das Speicherlevel der betrachteten Stunde muss die Summe sein aus dem Level der vorherhigen Stunde + Einspeicherung - Ausspeicherung
@constraint(model, Speicherstand_2[s in s_set, l in l_set], y[1,s,l] == y[t,s,l] + x[t,l,s] - x[t,s,l]) # NB 5 sagt, dass das Speicherlevel zu Stunde 1 gleich dem Speicherstand der letzten betrachteten Stunde sein muss
@constraint(model, Speicherstand_3[s in s_set, l in l_set], y[1,s,l] == 0.5*Volumenfaktor[j][s][l] * Kapazität[j][l][s]) # NB 5 sagt, dass das Speicherlevel zur Stunde Null der halben Kapazität entsprechen muss
@constraint(model, Emissionen[t in t_set, l in l_set], z[t,l] == sum(x[t,k,l] * Emissionsfaktor_el[j][k] for k in k_set))

@constraint(model, Speicherstand_1W[t in t_set[2:end], l in l_set], sw[t,l] == sw[t-1,l] - x[t-1,spw,l] + 0.0006*Volumenfaktor[j][spw][l] * Kapazität[j][l][spw]) # NB 4 gibt das Speicherlevel aus. Das Speicherlevel der betrachteten Stunde muss die Summe sein aus dem Level der vorherhigen Stunde + Einspeicherung - Ausspeicherung
@constraint(model, Speicherstand_2W[l in l_set], sw[1,l] == sw[t,l] - x[t,spw,l] + 0.0006*Volumenfaktor[j][spw][l] * Kapazität[j][l][spw]) # NB 5 sagt, dass das Speicherlevel zu Stunde 1 gleich dem Speicherstand der letzten betrachteten Stunde sein muss
@constraint(model, Speicherstand_3W[l in l_set], sw[1,l] == 0.5*Volumenfaktor[j][spw][l] * Kapazität[j][l][spw]) # NB 5 sagt, dass das Speicherlevel zur Stunde Null der halben Kapazität entsprechen muss

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

Nachfrage_df
Nachfrage_j = filter(row -> row.Jahr == j, Nachfrage_df)
Nachfrage_t = Nachfrage_j[1:t,:]

# Vorbereitung des Ergebnis-Outputs in Excel
# Dataframe für jedes Land wird erstellt, den Ergebnissen von oben werden die Exporte in die verschiedenen Länder angehängt, sowie die Einspeicherung und die Nachfrage des betrachteten Ladens  

function l_df1(l)
l1 = DataFrame(Array(x_results[:,:,l]), Ueberschriften)
return l1
end


function l_df2(l)
l2 = DataFrame(
    hcat(DE_df[:,l*"_im"], FR_df[:,l*"_im"], NL_df[:,l*"_im"], PL_df[:,l*"_im"], SE_df[:,l*"_im"], NO_df[:,l*"_im"], AT_df[:,l*"_im"], Pumpspeicher[:,l], PS_Speicherstand[:,l], Batteriespeicher[:,l], BS_Speicherstand[:,l], Wasserstoffspeicher[:,l], WS_Speicherstand[:,l], Nachfrage_t[:,l], Strompreise[:,l], Emissionen[:,l]),
    ["DE_ex", "FR_ex", "NL_ex", "PL_ex", "SE_ex", "NO_ex", "AT_ex", "PS_Einspeicherung", "PS_Speicherstand", "BS_Einspeicherung", "BS_Speicherstand", "WS_Einspeicherung", "WS_Speicherstand", "Nachfrage", "Strompreis", "Emissionen"])
return l2
end

Ergebnisse = []
for l in l_set
    l1 = l_df1(l)
    l2 = l_df2(l)
    le = hcat(l1, l2)
    push!(Ergebnisse, le)
end

# Export der vorbereiteten DataFrames in Excel
# Die Namen der Tabellenblätter müssen händisch erweitert werden, falls Länder & Speicher hinzugefügt werden
XLSX.writetable("Ergebnisse$j.xlsx", overwrite=true, 
        "DE" => Ergebnisse[1],
        "FR" => Ergebnisse[2],
        "NL" => Ergebnisse[3],
        "PL" => Ergebnisse[4],
        "SE" => Ergebnisse[5],
        "NO" => Ergebnisse[6],
        "AT" => Ergebnisse[7],
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
end