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


Zuordnung = Dict(
    :Wirkungsgrade => Dict(
        Kraftwerke_df[:,:Kategorie] .=> Kraftwerke_df[:,:Wirkungsgrad]
    ),
    :Brennstoffe => Dict(
        Kraftwerke_df[:,:Kategorie] .=> Kraftwerke_df[:,:Energieträger]
    ),
    :Brennstoffkosten => Dict(
        Energieträger_df[:, :Energieträger] .=> Energieträger_df[:,:Brennstoffkosten]
    ),
    :Emissionsfaktor => Dict(
        Energieträger_df[:, :Energieträger] .=> Energieträger_df[:, :Emissionsfaktor]
    )
)


Kapazität = Dict(
        "DE" => Dict(Kraftwerke_df[:,:Kategorie] .=> Kapazität_df[:,:DE]),
        "FR" => Dict(Kraftwerke_df[:,:Kategorie] .=> Kapazität_df[:,:FR]), 
        "NL" => Dict(Kraftwerke_df[:,:Kategorie] .=> Kapazität_df[:,:NL]) 
        )


        using CSV

function arr_to_csv(x, outputstring)
    df = DataFrame(t = Int[], k = Int[], l = Int[], x_results = Float64[])
    sizes = size(x)

    for t in 1:sizes[1]
        for k in 1:sizes[2]
            for l in 1:sizes[3]
                    push!(df, (t, k, l, x_results[t,k,l]))
            end
        end
    end
    df |> CSV.write(outputstring, header = ["t", "k", "l", "value"])
end
arr_to_csv(x_results, "test.csv")


"Braunkohle_0" => Dict(
            "DE" => Dict(t_set .=> fossil),
            "FR" => Dict(t_set .=> fossil),
            "NL" => Dict(t_set .=> fossil)),
        "Braunkohle_-" => Dict(
            "DE" => Dict(t_set .=> fossil),
            "FR" => Dict(t_set .=> fossil),
            "NL" => Dict(t_set .=> fossil)),
        "Braunkohle_+" => Dict(
            "DE" => Dict(t_set .=> fossil),
            "FR" => Dict(t_set .=> fossil),
            "NL" => Dict(t_set .=> fossil)),
        "Steinkohle_0" => Dict(
            "DE" => Dict(t_set .=> fossil),
            "FR" => Dict(t_set .=> fossil),
            "NL" => Dict(t_set .=> fossil)),
        "Steinkohle_-" => Dict(
            "DE" => Dict(t_set .=> fossil),
            "FR" => Dict(t_set .=> fossil),
            "NL" => Dict(t_set .=> fossil)),
        "Steinkohle_+" => Dict(
            "DE" => Dict(t_set .=> fossil),
            "FR" => Dict(t_set .=> fossil),
            "NL" => Dict(t_set .=> fossil)),
        "Erdgas_0" => Dict(
            "DE" => Dict(t_set .=> fossil),
            "FR" => Dict(t_set .=> fossil),
            "NL" => Dict(t_set .=> fossil)),
        "Erdgas_-" => Dict(
            "DE" => Dict(t_set .=> fossil),
            "FR" => Dict(t_set .=> fossil),
            "NL" => Dict(t_set .=> fossil)),
        "Erdgas_+" => Dict(
            "DE" => Dict(t_set .=> fossil),
            "FR" => Dict(t_set .=> fossil),
            "NL" => Dict(t_set .=> fossil)),
        "Wasserkraft" => Dict(
            "DE" => Dict(t_set .=> fossil),
            "FR" => Dict(t_set .=> fossil),
            "NL" => Dict(t_set .=> fossil)),
        "Biomasse" => Dict(
            "DE" => Dict(t_set .=> fossil),
            "FR" => Dict(t_set .=> fossil),
            "NL" => Dict(t_set .=> fossil)),
        "Windkraft" => Dict(
            "DE" => Dict(t_set .=> wind("DE")),
            "FR" => Dict(t_set .=> wind("FR")),
            "NL" => Dict(t_set .=> wind("NL"))),
        "PV" => Dict(
            "DE" => Dict(t_set .=> sonne("DE")),
            "FR" => Dict(t_set .=> sonne("FR")),
            "NL" => Dict(t_set .=> sonne("NL")))
            


# Einzelne Excellisten für die verschiedenen Länder und deren Ergebnisse erstellen
for z in l_set
    y_results = x_results[:,:,z]
    z_results = Array(y_results)
    Excelname = "Ergebnisse"*z*".xlsx"
    results = DataFrame(z_results, k_set)
    rm(Excelname, force=true) #Lösche die alte, bereits bestehende Excel-Ergebnisliste
    XLSX.writetable(Excelname, results) #Erstelle eine neue Ergebnisliste
end


# Inhalte fehlen noch 
XLSX.openxlsx("test_file.xlsx", mode="w") do xf
    XLSX.rename!(xf[1], "first")

    for sheetname in l_set
      XLSX.addsheet!(xf, sheetname)
    end

    for q in 2:l+1 
        xf[q]["A1"] = "A"
    end
    
end