## expefX Funktion

### Beskrivning
Denna funktion exporterar de markerade kalkylbladen som PDF-filer till den plats där arbetsboken är sparad, med filnamn baserat på värdet i cell O1.

### Användning
1. Ange det önskade filnamnet i cell O1.
2. Använd funktionen "expefX()" för att köra exporten.
3. De markerade kalkylbladen ("sign.alla" och "sign.behandlad") kommer att exporteras som PDF-filer till platsen där arbetsboken är sparad.
4. En dialogruta visas när exporten är klar.

### Anmärkningar
- Funktionen kopierar värdet från cell O2 till O1 och använder sedan värdet i cell O1 som filnamn för PDF-filerna.
- Filnamnet för PDF-filerna skapas baserat på värdet i cell O1, med eventuella kolon (:) ersatta med understreck (_).
- PDF-filerna exporteras till samma plats som arbetsboken är sparad i.
- Om det redan finns en fil med samma namn som det angivna filnamnet kommer den att ersättas.
- Funktionen antar att kalkylbladen "sign.alla" och "sign.behandlad" finns i arbetsboken.
