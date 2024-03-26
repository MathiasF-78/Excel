## SplitWorksheet Macro

### Beskrivning
Detta makro delar en arbetsbok i flera arbetsblad genom att flytta ett angivet antal rader till varje nytt arbetsblad.

### Användning
1. Kör makrot "SplitWorksheet".
2. Ange det önskade antalet rader som ska finnas på varje nya arbetsblad när du körs.
3. Makrot kommer att skapa nya arbetsblad och flytta rader från det aktiva arbetsbladet till de nya arbetsbladen tills det inte finns fler rader kvar att flytta.

### Anmärkningar
- Makrot använder det angivna antalet rader för att bestämma när det ska skapas nya arbetsblad och när rader ska flyttas.
- Ett nytt arbetsblad skapas för varje grupp av rader som flyttas från det aktiva arbetsbladet.
- Namnen på de nya arbetsbladen får namn baserat på det aktiva arbetsbladets namn följt av en räknare för att ange ordningen på de nya arbetsbladen.
- Se till att spara arbetsboken innan du kör makrot, så att du kan återställa den om det behövs.
