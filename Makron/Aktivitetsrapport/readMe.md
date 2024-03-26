## Aktivitetsrapport Macro

### Beskrivning
Denna modul innehåller makron som utför olika åtgärder relaterade till hantering av data och formatering av kalkylblad i Excel.

### Makron

#### everything
- **Beskrivning:** Detta makro utför flera åtgärder inklusive att formatera rubriker, kontrollera och lägga till eventuella rubriker, samt summera och beräkna procentandelar.
- **Användning:** Kör makrot för att utföra de nämnda åtgärderna.
- **Anmärkningar:** Makrot använder andra undermakron för att utföra specifika uppgifter.

#### CheckColumnHeadingIsPresent
- **Beskrivning:** Detta makro kontrollerar om en specifik kolumnrubrik finns i det aktiva kalkylbladet och lägger till den om den saknas.
- **Användning:** Används internt av makrot "everything" för att säkerställa att viktiga kolumnrubriker finns.

### Anmärkningar
- Modulen innehåller även andra makron som används internt av huvudmakrot "everything".
- Makrona i modulen är utformade för att underlätta hantering och formatering av data i Excel-kalkylblad.
