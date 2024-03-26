## BFUS_orgnr Macro

### Beskrivning
Detta makro formaterar organisationsnummer genom att ta bort de första två tecknen om de är "16" och de sista två tecknen om de är "00". Till exempel, om organisationsnumret är "16xxxxxxxxxx00", kommer makrot att ändra det till "xxxxxxxxxx".

### Användning
1. Placera organisationsnumren som du vill formatera i kolumn A i arket.
2. Kör makrot "BFUS_orgnr".

### Anmärkningar
- Makrot förutsätter att organisationsnumren finns i kolumn A, från rad 1 till den sista raden med data.
- Endast organisationsnummer som följer mönstret "16xxxxxxxxxx00" kommer att formateras av makrot.
- Makrot bearbetar organisationsnummer en rad i taget och tar bort "16" från början och "00" från slutet av varje organisationsnummer i kolumn A.
