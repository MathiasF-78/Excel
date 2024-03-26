
# SendAsMail Makro för Excel

Detta är ett VBA-makro för Microsoft Excel och Outlook som möjliggör automatisk sändning av det aktiva Excel-dokumentet som en e-postbilaga. Makrot använder Outlook för att skapa och skicka e-postmeddelandet, komplett med en HTML-formaterad kropp och en signatur hämtad från användarens Outlook-inställningar.

## Funktionalitet

- **Skicka Aktuellt Dokument:** Automatiskt bifogar och skickar det aktuellt öppna Excel-dokumentet.
- **Anpassningsbart Meddelande:** Möjlighet att anpassa e-postmeddelandets text och ämnesrad.
- **Automatisk Signatur:** Inkluderar användarens Outlook-signatur i e-postmeddelandet.

## Användning

För att använda detta makro, kopiera hela modulen till ditt VBA-projekt i Outlook. Konstaterat fungerar med Office 2016.

### Steg för Anpassning

1. **Meddelandetext:** Ändra `"Hej,"` och `"Se bifogat dokument."` i `strbody` variabeln för att anpassa e-postmeddelandets inledande hälsning och huvudtext.
2. **Signaturfil:** Byt ut `"Mathias.htm"` med namnet på din egen signaturfil. Din signaturfilens namn brukar oftast vara `"Standard.htm"` men kan variera.
3. **Mottagare och Ämnesrad:** Ange standardmottagare i `.To`, `.CC`, och `.BCC` samt anpassa ämnesraden i `.Subject` enligt önskemål.

### Skicka eller Visa Förhandsgranskning

- Använd `.Send` metoden för att skicka e-postmeddelandet direkt.
- Använd `.Display` metoden istället om du vill visa en förhandsgranskning av e-postmeddelandet innan det skickas.

## Noteringar

Detta makro använder Outlooks objektmodell och kräver att Microsoft Outlook är installerat och konfigurerat på användarens dator.

## Licens

Detta makro är tillgängligt för användning och modifiering utan restriktioner.
