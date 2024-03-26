## SheetAsQuote Macro

### Beskrivning
Detta makro exporterar den aktiva kalkylbladet som en PDF-fil med ett specificerat filnamn baserat på cellinnehållet.

### Användning
1. Placera ett ActiveX-kommandoknappskontroll (CommandButton) på det aktiva kalkylbladet i Excel.
2. Kopiera och klistra in makrokoden i modulen för det aktiva kalkylbladet.
3. När användaren klickar på kommandoknappen kommer makrot att köras och exportera kalkylbladet som en PDF-fil med det angivna filnamnet.

### Anmärkningar
- Makrot förutsätter att det finns en kommandoknapp på det aktiva kalkylbladet och att användaren klickar på den för att starta exportprocessen.
- Filnamnet för den exporterade PDF-filen genereras baserat på värdet i cellen F8 och det aktuella datumet.
- Det är viktigt att notera att PDF-filen sparas i samma mapp som den aktuella arbetsboken och öppnas automatiskt efter att den har skapats.
- Om du vill anpassa filnamnet eller ytterligare anpassningar i exportprocessen kan du redigera makrokoden efter behov.
