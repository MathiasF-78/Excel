## Speak Macro

### Beskrivning
Detta makro spelar upp en rolig meddelande för användaren och simulerar en självförstöring av CRM-systemet med en nedräkning från 10 till 8 och sedan en meddelande som säger "Just kidding!".

### Användning
1. Kör makrot "Speak".
2. Ett meddelande visas som simulerar en självförstöring av CRM-systemet med en nedräkning.
3. Efter nedräkningen spelar makrot upp meddelandet "Just kidding!".

### Anmärkningar
- Om användaren anger en text i kommenterad kod som används för nedräkningen, kommer den texten att ersättas med det angivna meddelandet.
- Makrot använder Windows-röstigenkänning (sapi.spvoice) för att läsa upp texten.
- Om ingen text anges vid körning av makrot, visas ett meddelande om att ingen text angavs och makrot avslutas.
