## ExporteraAllaModuler Macro

### Beskrivning
Detta makro exporterar alla moduler (kodmoduler, formulär och klassmoduler) från det aktuella Excel-projektet och sparar dem som separata filer på C:\.

### Användning
1. Kör makrot.
2. Alla moduler i det aktuella Excel-projektet kommer att exporteras som separata filer och sparas i roten på C:\.
3. Filerna kommer att namnges enligt deras respektive modultyper: ".bas" för kodmoduler, ".frm" för formulär och ".cls" för klassmoduler.

### Anmärkningar
- Om filerna redan finns på C:\ kommer de att ersättas utan varning.
- Makrot tar inte hänsyn till eventuella submappar eller filstrukturer på C:\ utan sparar alla filer i roten på enheten.
- Var noga med att säkerhetskopiera dina filer innan du kör detta makro, särskilt om du har andra filer sparade på C:\.
