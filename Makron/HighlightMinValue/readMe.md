## HighlightMinValue Macro

### Beskrivning
Detta makro markerar det lägsta värdet i det markerade cellområdet i Excel.

### Användning
1. Markera det cellområde där du vill markera det lägsta värdet.
2. Kör makrot "HighlightMinValue".

### Åtgärder som utförs av makrot:
- Makrot går igenom varje cell i det markerade området och jämför värdet med det lägsta värdet i området.
- Om en cell har det lägsta värdet, markeras den med en särskild stil (i detta fall "Good").

### Anmärkningar
- Endast det första förekommandet av det lägsta värdet i det markerade området markeras.
- Om flera celler har samma lägsta värde, markeras endast den första cellen med detta värde.
- Makrot antar att det lägsta värdet är unikt i det markerade området.
- Om det inte finns några numeriska värden i det markerade området eller om alla celler har samma värde, kommer inga celler att markeras.

