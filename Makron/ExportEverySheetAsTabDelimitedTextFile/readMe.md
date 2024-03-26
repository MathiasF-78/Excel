## Export_Each_Sheet_As_TabDelimited_TextFile Macro

### Beskrivning
Detta makro exporterar varje kalkylblad i arbetsboken som en textfil med tabbavgränsad formatering och sparar dem i en separat mapp.

### Användning
1. Kör makrot.
2. Varje kalkylblad kopieras och sparas som en separat textfil i en mapp med namnet "[datum]_ArbetsbokensNamn" i samma mapp som den aktuella arbetsboken.
3. Textfilerna kommer att namnges efter respektive kalkylblad.

### Anmärkningar
- Makrot skapar en ny mapp med namnet "[datum]_ArbetsbokensNamn" i samma mapp som den aktuella arbetsboken, om mappen inte redan finns.
- Varje textfil kommer att innehålla data från det motsvarande kalkylbladet, med tabbavgränsad formatering.
- Om du vill inkludera en tidsstämpel i mappnamnet, avkommentera raden som använder variabeln MyStr2 för tidsstämpeln.
