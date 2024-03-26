## AddTimestamp Macro

### Beskrivning
Detta makro infogar en tidsstämpel i den intilliggande cellen när detekterar en ändring i en cell i kolumn A.

### Användning
1. Makrot aktiveras automatiskt när en ändring görs i en cell i kolumn A.
2. Om en cell i kolumn A ändras och har ett värde, infogas en tidsstämpel i den intilliggande cellen i kolumn B.

### Anmärkningar
- Tidsstämpeln visas i formatet "dd-mm-yyyy hh:mm:ss".
- Om en cell i kolumn A redan har ett värde och sedan ändras, uppdateras tidsstämpeln i den intilliggande cellen.
- Om makrot upptäcker något fel under körningen, kommer det att fortsätta utan att visa något felmeddelande.
