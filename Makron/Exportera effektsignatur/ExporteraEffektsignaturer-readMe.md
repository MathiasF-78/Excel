## ExporteraEffektsignaturer Funktion

### Beskrivning
Denna funktion exporterar valda kalkylblad till PDF-format och sparar dem i samma mapp som den aktuella arbetsboken.

### Användning
1. Kör funktionen.
2. Kalkylbladen "sign.alla" och "sign.behandlad" väljs automatiskt för export till PDF.
3. PDF-filerna sparas i samma mapp som den aktuella arbetsboken med namnen "O1_värde.pdf", där "O1_värde" ersätts med värdet i cellen O1.

### Anmärkningar
- Funktionen tar värdet från cellen O1 som en del av PDF-filnamnet.
- Om PDF-filerna redan finns i mappen kommer de att ersättas.
- Ett meddelande visas när alla PDF-filer har exporterats framgångsrikt.
