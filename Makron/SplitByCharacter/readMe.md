# SplitCell Macro

## Beskrivning
Detta makro delar upp data i valda celler i två delar baserat på ett angivet tecken. Den extraherade datan kommer att antingen skrivas direkt till höger om det befintliga värdet eller i en ny kolumn beroende på användarens val.

## Användning
1. Markera de celler som innehåller data som du vill dela upp.
2. Kör makrot.
3. Ange tecknet som ska användas för att dela upp cellerna.
4. Om du vill skapa en ny kolumn innan data extraheras, klicka "Ja" när du tillfrågas. Annars klicka "Nej".
5. Makrot kommer att dela upp cellerna och skriva den extraherade datan i antingen den befintliga kolumnen eller i en ny kolumn beroende på ditt val.

## Exempel
Om du har celler med textvärden som innehåller ett mellanslag och du vill dela upp dem i förnamn och efternamn, kan du använda detta makro för att göra det automatiskt.

## Anmärkningar
- Om tecknet inte finns i en cell, kommer en varning att visas för användaren och ingen åtgärd utförs för den cellen.
- Se till att markera de celler som innehåller data som du vill dela upp innan du kör makrot.
- Om du väljer att skapa en ny kolumn innan data extraheras, se till att det inte finns befintlig data i den nya kolumnen för att undvika oönskade resultat.
