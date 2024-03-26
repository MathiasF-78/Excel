## FileNameAsCellContent Macro

### Beskrivning
Detta makro sparar den aktuella arbetsboken med ett filnamn som hämtas från innehållet i en specifik cell.

### Användning
1. Kör makrot.
2. En dialogruta för att välja sparplats kommer inte att visas eftersom sökvägen är förinställd i koden. Du kan ändra sökvägen i koden enligt behov.
3. Filnamnet för den sparade arbetsboken kommer att vara samma som innehållet i cell A1 i samma arbetsbok, med filändelsen .xlsx.

### Anmärkningar
- Innan du kör makrot, se till att cellen A1 innehåller önskat filnamn.
- Ändra sökvägen i variabeln "Path" enligt dina önskemål innan du kör makrot.
- Makrot sparar arbetsboken som en .xlsx-fil. Om du vill ändra filformatet, ändra xlOpenXMLWorkbook till det önskade formatet enligt länken: [Filformat för Excel](http://msdn.microsoft.com/en-us/library/office/ff198017.aspx).
- Efter att arbetsboken har sparats kommer den att automatiskt stängas.
