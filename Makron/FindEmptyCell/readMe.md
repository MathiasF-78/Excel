## FindEmptyCell Macro

### Beskrivning
Detta makro letar efter den första tomma cellen nedanför den aktiva cellen i kolumnen.

### Användning
1. Placera den aktiva cellen i den kolumn där du vill börja söka efter tomma celler.
2. Kör makrot.
3. Makrot kommer att hitta den första tomma cellen nedanför den aktiva cellen och placera markören där.

### Anmärkningar
- Makrot söker efter tomma celler vertikalt nedåt från den aktiva cellen i den aktuella kolumnen.
- Om den aktiva cellen redan är tom kommer makrot att flytta markören till nästa cell och börja söka därifrån.
- Om det inte finns några tomma celler under den aktiva cellen kommer markören att placeras i den sista cellen i kolumnen.
