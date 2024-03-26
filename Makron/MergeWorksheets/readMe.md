## MergeWorksheets Macro

### Beskrivning
Detta makro används för att sammanfoga data från flera arbetsblad i en arbetsbok till ett enda arbetsblad kallat "Master".

### Användning
1. Kör makrot "MergeWorksheets".
2. Data från alla arbetsblad förutom "Master" kommer att sammanfogas i det nya "Master"-arbetsschemat.

### Åtgärder som utförs av makrot:
- Skapar ett nytt arbetsblad med namnet "Master" (om det inte redan finns ett arbetsblad med detta namn).
- Kopierar kolumnrubrikerna från det första arbetsbladet till "Master"-arbetsschemat.
- Sammanfogar data från alla arbetsblad (förutom "Master") under kolumnrubrikerna på "Master"-arbetsschemat.
- Justerar kolumnbredden för att passa den sammanfogade datan.

### Anmärkningar
- Om det redan finns ett arbetsblad med namnet "Master" i arbetsboken kommer en varningsruta att visas, och sammanfogningen kommer inte att utföras tills "Master"-arbetsschemat har blivit borttaget eller omdöpt.
- Data från de första 65 536 raderna (maximalt antal rader i ett Excel-ark) av varje arbetsblad kommer att sammanfogas.
- Kolumnbredden på "Master"-arbetsschemat kommer att justeras baserat på den sammanfogade datan.

