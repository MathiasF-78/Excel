## ExportVisualBasicCode Macro

### Beskrivning
Detta makro exporterar all VBA-källkod i det aktuella Excel-projektet till separata textfiler för att underlätta versionshantering med korrekt källkodsversionskontroll.

### Användning
1. Kör makrot.
2. All VBA-källkod i projektet exporteras som separata textfiler och sparas i en mapp med namnet "VisualBasic" i samma mapp som den aktuella arbetsboken.
3. Textfilerna namnges enligt namnet på VBA-komponenten (modul, klassmodul, formulär osv.) med rätt filändelse (.bas, .cls, .frm).

### Anmärkningar
- Makrot kräver att tillåta åtkomst till VBA-projektobjektmodellen under "Inställningar för makron" i Excel för att fungera korrekt.
- Om mappen "VisualBasic" inte finns skapas den automatiskt.
- Om det uppstår fel vid export av någon VBA-komponent visas ett felmeddelande och exporten avbryts för den specifika komponenten.
- Efter framgångsrik export visas en statusindikator längst ner i Excel som meddelar antalet exporterade VBA-filer och sökvägen till mappen där de sparades.
