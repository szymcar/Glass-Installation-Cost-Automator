let
    // 1. Konfiguracja ścieżki dynamicznej z tabeli parametrów w Excelu
    TabelaWExcelu = Excel.CurrentWorkbook(){[Name="Tabela_Sciezka"]}[Content],
    Folder = TabelaWExcelu{0}[Sciezka],    
    
    // 2. Budowanie ścieżki i ładowanie pliku źródłowego
    PelnaSciezka = if Text.EndsWith(Folder, "\") then Folder & "zamGlassRysMasaCena.csv" else Folder & "\zamGlassRysMasaCena.csv",
    Źródło = Csv.Document(File.Contents(PelnaSciezka),[Delimiter=";", Columns=26, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    
    // 3. Dynamiczne znajdowanie nagłówka tabeli (pominięcie zbędnych wierszy raportu)
    PominiecieWierszy = Table.Skip(Źródło, each ([Column1] <> "L.p.")),
    Naglowki = Table.PromoteHeaders(PominiecieWierszy, [PromoteAllScalars=true]),
    
    // 4. Selekcja kolumn i czyszczenie typów danych (przygotowanie do obliczeń)
    WybraneKolumny = Table.SelectColumns(Naglowki,{"L.p.","Ilość", "Masa[kg] razem"}),
    TypyDanych = Table.TransformColumnTypes(WybraneKolumny, {{"Masa[kg] razem", type number}, {"Ilość", Int64.Type}, {"L.p.", Int64.Type}}),
    BezBledow = Table.RemoveRowsWithErrors(TypyDanych),
    BezPustych = Table.SelectRows(BezBledow, each ([#"L.p."] <> null)),
    
    // 5. Wyliczanie masy jednostkowej (Masa Razem / Ilość)
    MasaJednostkowa = Table.AddColumn(BezPustych, "Masa pojedynczej", each [#"Masa[kg] razem"] / [Ilość], type number),
    
    // 6. Logika biznesowa: Przypisanie czasu pracy (h) w zależności od wagi elementu
    // Uwaga: Progi czasowe wyznaczone nieliniowo na podstawie wagi pojedynczej tafli
    CzasPracy = Table.AddColumn(MasaJednostkowa, "Całkowity czas [h]", each 
        if [Masa pojedynczej] > 500 then [Ilość] * 1 // 1h na sztukę
        else if [Masa pojedynczej] > 400 then [Ilość] * (1/3) // 20 min na sztukę
        else if [Masa pojedynczej] > 300 then [Ilość] * (1/5) // 12 min na sztukę
        else if [Masa pojedynczej] > 200 then [Ilość] * (1/8) // 7.5 min na sztukę
        else if [Masa pojedynczej] > 150 then [Ilość] * (1/10) // 6 min na sztukę
        else 0, type number),

    // 7. Agregacja końcowa czasu zaszklenia dla całego zlecenia
    SumaCzasu = List.Sum(CzasPracy[#"Całkowity czas [h]"])
in
    SumaCzasu