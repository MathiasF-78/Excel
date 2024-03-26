## FilnamnCellvarde Macro

### Beskrivning
Detta makro sparar den aktuella arbetsboken med ett filnamn som motsvarar innehållet i cell A1 i samma arbetsbok.

### Användning
1. Se till att cellen A1 innehåller det önskade filnamnet.
2. Kör makrot.
3. Arbetsboken sparas med filnamnet som motsvarar innehållet i cell A1 i den angivna mappen "C:".

### Anmärkningar
- Makrot ersätter inte befintliga filer med samma namn utan att ge något varningsmeddelande.
- Om du vill ändra sökvägen där filen sparas, ändra värdet för variabeln "FPath" i koden.
- Se till att det finns giltigt filnamn i cellen A1 innan du kör makrot, annars kommer det att uppstå fel.
