## neweffektexport Macro

### Beskrivning
Detta makro kopierar värdet från cellen O2 och klistrar in det i cellen O1 som enbart värde. Dessutom finns en funktion (SaveSelectToPDFv6) för att exportera valda kalkylblad till PDF-filer.

### Användning
1. För att kopiera värdet från cellen O2 och klistra in det i cellen O1 som enbart värde, kör makrot "neweffektexport".
2. För att exportera valda kalkylblad till PDF-filer, använd funktionen "SaveSelectToPDFv6".

### Anmärkningar
- Makrot "neweffektexport" antar att det aktiva arket är det arket där kopieringen och klistringen ska ske.
- Funktionen "SaveSelectToPDFv6" exporterar de valda kalkylbladen till PDF-filer och sparar dem i den aktuella arbetsbokens mapp. Filnamnen genereras från värdet i cellen O1 och eventuella kolon (":") i detta värde ersätts med understreck ("_").
- Det är viktigt att notera att funktionen "SaveSelectToPDFv6" är konfigurerad för att exportera kalkylbladen "sign.alla" och "sign.behandlad" till PDF-filer. Om andra kalkylblad ska exporteras måste koden anpassas därefter.
