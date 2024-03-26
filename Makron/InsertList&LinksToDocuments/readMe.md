## InsertList Macro

### Beskrivning
Detta makro används för att infoga hyperlänkar till filer i en specifik mapp i Excel. Varje fil i mappen kommer att vara länkad till en separat cell i kolumn A, med filnamnet som text.

### Användning
1. Se till att mappen med filerna du vill länka till är tillgänglig.
2. Kör makrot "InsertList".

### Åtgärder som utförs av makrot:
- Makrot använder en instans av FileSystemObject för att komma åt filsystemet.
- Det går igenom varje fil i den angivna mappen och skapar en hyperlänk i kolumn A för varje fil.
- Varje hyperlänk är länkad till den faktiska filen och använder filens namn som text för hyperlänken.

### Anmärkningar
- Endast filer i den angivna mappen kommer att vara länkade i Excel-arket.
- Varje fil i mappen genererar en separat hyperlänk i Excel, med filens namn som text för hyperlänken.
- Om det finns många filer i mappen kan resultatet bli en lång lista med hyperlänkar i kolumn A.

