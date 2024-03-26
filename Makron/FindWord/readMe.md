## FindWord och FindWordRev Funktioner

### Beskrivning
Dessa funktioner används för att hitta ord från en textsträng baserat på deras position i strängen. Funktionen "FindWord" letar efter ett ord i strängen från början, medan funktionen "FindWordRev" letar efter ett ord från slutet av strängen.

### Användning
#### FindWord Funktionen
- Syntax: `FindWord(Source As String, Position As Integer) As String`
- Parametrar:
  - Source: Den ursprungliga textsträngen där sökningen ska utföras.
  - Position: Positionen för det önskade ordet i strängen.
- Returnerar: Det önskade ordet från den angivna positionen i strängen.
- Exempel: `FindWord("Detta är en exempeltext", 3)` kommer att returnera ordet "en".

#### FindWordRev Funktionen
- Syntax: `FindWordRev(Source As String, Position As Integer) As String`
- Parametrar:
  - Source: Den ursprungliga textsträngen där sökningen ska utföras.
  - Position: Positionen för det önskade ordet från slutet av strängen.
- Returnerar: Det önskade ordet från den angivna positionen i strängen räknat från slutet.
- Exempel: `FindWordRev("Detta är en exempeltext", 2)` kommer att returnera ordet "en".

### Anmärkningar
- Funktionerna delar upp textsträngen i ord och returnerar det ord som matchar den angivna positionen.
- Om den angivna positionen är utanför det tillgängliga antalet ord i strängen returneras en tom sträng.
- Tomma utrymmen och skiljetecken mellan orden hanteras av funktionerna för att returnera ordet korrekt.
