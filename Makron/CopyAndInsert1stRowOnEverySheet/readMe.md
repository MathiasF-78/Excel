## CopyToAllSheets Macro

### Beskrivning
Detta makro kopierar data från en rad till alla andra ark i arbetsboken utom ett specifikt ark. Det används vanligtvis när du vill sprida samma data över flera ark i arbetsboken för att hålla dem synkroniserade.

### Användning
1. Se till att ha ett ark i arbetsboken med namnet "deltagare". Detta ark kommer att fungera som källan för datakopiering.
2. Kör makrot "CopyToAllSheets".
3. Makrot kopierar dataraden från arket "deltagare" till den första raden på alla andra ark i arbetsboken.

### Anmärkningar
- Makrot förutsätter att det finns minst en rad med data på arket "deltagare".
- Om det finns några tomma rader innanför den första raden på de andra ark som du vill kopiera till, kommer makrot att infoga en ny rad för att placera data i den första raden på dessa ark.
- Se till att ange rätt namn på källarket ("deltagare") och att det finns minst en rad med data på detta ark innan du kör makrot.
