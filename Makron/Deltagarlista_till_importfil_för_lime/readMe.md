## ImportDeltagareToLime Macro

### Beskrivning
Denna makro delar upp en arbetsbok i flera arbetsblad, formaterar dem och exporterar dem som tabbavgränsade textfiler.

### Steg
1. **Delar upp arbetsblad:** Delar upp huvudarbetsbladet i flera arbetsblad med högst 1000 rader var. Varje nytt arbetsblad kommer att namnges som "Huvudarknamn (x)", där "x" är ett nummer.
2. **Kopierar första raden:** Kopierar första raden från huvudarkivet och lägger till den i varje nytt ark.
3. **Exporterar som textfiler:** Exporterar varje ark som en tabbavgränsad textfil till en specifik mapp med dagens datum som en del av mappnamnet.

### Anmärkningar
- Varje nytt ark får en kopia av den första raden från huvudarket.
- Arbetsbladen delas upp i mindre delar om de innehåller fler än 1000 rader.
- De exporterade textfilerna sparas i en mapp med dagens datum som en del av mappnamnet.
