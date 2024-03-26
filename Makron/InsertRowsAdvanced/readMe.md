## InsertRowsAdvanced Macro

### Beskrivning
Detta makro används för att infoga flera rader med ett visst mellanrum mellan varje insättning i Excel.

### Användning
1. Kör makrot "InsertRowsAdvanced".
2. Ange antalet rader du vill infoga.
3. Ange antalet rader du vill hoppa över mellan varje insättning.

### Åtgärder som utförs av makrot:
- Makrot frågar användaren hur många rader de vill infoga och hur många rader de vill hoppa över mellan varje insättning.
- Det markerar sedan den aktuella cellen och börjar infoga rader med det angivna mellanrummet och antalet rader tills den når en tom cell.
- Varje insättning sker med det angivna mellanrummet mellan raderna.

### Anmärkningar
- Endast tomma rader infogas av makrot.
- Om en cell innehåller data kommer makrot att sluta infoga rader och avslutas.
- Makrot antar att användaren anger ett positivt heltal för både antalet rader som ska infogas och antalet rader som ska hoppas över.

