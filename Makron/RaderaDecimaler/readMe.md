## RaderaDecimaler Macro

### Beskrivning
Detta makro kopierar och avrundar värdena i de markerade cellerna till närmaste heltal och klistrar sedan in dem utan decimaler.

### Användning
1. Markera det cellområde där du vill ta bort decimalerna från värdena.
2. Kör makrot "RaderaDecimaler".

### Anmärkningar
- Makrot förutsätter att du markerar det cellområde där värdena ska ändras.
- Endast numeriska värden i de markerade cellerna kommer att påverkas av makrot.
- Makrot använder Excel-funktionen "Round" för att avrunda värdena till närmaste heltal och sedan klistrar in dem utan decimaler med hjälp av "PasteSpecial" -funktionen.
- Det är viktigt att notera att detta makro inte påverkar celler som innehåller text eller formler.
