## sumByYear Macro

### Beskrivning
Detta makro summerar värden baserat på år i en given kolumn och skriver sedan summeringen bredvid respektive år.

### Användning
1. Kör makrot "sumByYear".
2. Makrot kommer att söka igenom en given kolumn (kolonn F) och summera värden från en annan kolumn (kolonn G) baserat på år.
3. Resultatet av varje summering (för varje år) skrivs bredvid respektive år i en ny kolumn (kolonn H för år och kolonn I för summeringen).

### Anmärkningar
- Makrot använder Excel-funktionen SUMIF för att summera värden baserat på ett givet kriterium (i detta fall år).
- Det förutsätts att åren är skrivna i formatet "ÅÅÅÅ" (till exempel "2015", "2016", osv.) i kolumn F och att värdena som ska summeras finns i kolumn G.
- Resultaten av summeringarna skrivs bredvid respektive år i kolumn H (för år) och kolumn I (för summeringen).
- Om det finns fler år att summera än de som är angivna i koden, måste koden uppdateras för att inkludera dessa år.
