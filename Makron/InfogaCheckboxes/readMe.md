## InfogaCheckboxes Macro

### Beskrivning
Detta makro används för att infoga kryssrutor i markerade celler i Excel och koppla dem till cellens värde.

### Användning
1. Markera de celler där du vill infoga kryssrutor.
2. Kör makrot "InfogaCheckboxes".

### Åtgärder som utförs av makrot:
- Makrot går igenom varje markerad cell och infogar en kryssruta i varje cell.
- Kryssrutornas läge och storlek anpassas till storleken på den markerade cellen.
- Varje kryssruta kopplas till cellen genom att den länkas till cellens adress.
- Om cellens värde är "x", kommer kryssrutan att vara markerad, annars kommer den vara tom.
- Format och färg på kryssrutorna ändras baserat på cellens värde.

### Anmärkningar
- Makrot antar att det finns en markerad cell där du vill infoga kryssrutor.
- Kryssrutornas utseende och koppling till cellvärden kan justeras genom att ändra parametrar och villkor i makrot.
- Felhantering är implementerad för att hantera eventuella fel under körningen av makrot.

