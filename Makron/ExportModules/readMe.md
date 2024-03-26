## ExportModules Macro

### Beskrivning
Detta makro exporterar alla kodmoduler, formulär och klassmoduler från det aktuella Excel-projektet och sparar dem som separata filer i en mapp.

### Användning
1. Kör makrot.
2. Alla kodmoduler, formulär och klassmoduler i det aktuella Excel-projektet exporteras som separata filer och sparas i en mapp.
3. Mappen skapas i "Mina dokument" och heter "VBAProjectFiles".
4. För att importera modulerna till en annan arbetsbok, använd makrot "ImportModules".

### Anmärkningar
- Om mappen "VBAProjectFiles" inte finns kommer den att skapas automatiskt.
- Om det inte finns några kodmoduler, formulär eller klassmoduler i det aktuella projektet visas ett meddelande.


## ImportModules Macro

### Beskrivning
Detta makro importerar kodmoduler, formulär och klassmoduler från en specifik mapp och lägger till dem i den aktuella Excel-arbetsboken.

### Användning
1. Kör makrot.
2. Alla kodmoduler, formulär och klassmoduler i mappen "VBAProjectFiles" importeras och läggs till i den aktuella Excel-arbetsboken.
3. Alla befintliga kodmoduler, formulär och klassmoduler i den aktuella arbetsboken kommer att raderas innan de nya modulerna importeras.

### Anmärkningar
- Om det inte finns några filer att importera från mappen "VBAProjectFiles" visas ett meddelande.
- Om den aktuella arbetsboken är skyddad med VBA-projektlösenord visas ett meddelande och importen avbryts.

