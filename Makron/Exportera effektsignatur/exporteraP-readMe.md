## exporteraP Macro

### Beskrivning
Detta makro exporterar de markerade kalkylbladen som PDF-filer till den plats där arbetsboken är sparad. Filnamnet för varje PDF-fil skapas genom att lägga till "Effektsignatur för" framför värdet i cell O1.

### Användning
1. Ange det önskade namnet på effektsignaturen i cell O1.
2. Kör makrot.
3. De markerade kalkylbladen ("sign.alla" och "sign.behandlad") kommer att exporteras som PDF-filer till platsen där arbetsboken är sparad.
4. En dialogruta visas när exporten är klar.

### Anmärkningar
- Makrot kopierar värdet från cell O2 till O1 och använder sedan värdet i cell O1 som en del av filnamnet för PDF-filerna.
- Filnamnet för varje PDF-fil skapas genom att lägga till "Effektsignatur för" framför värdet i cell O1, med eventuella kolon (:) ersatta med understreck (_).
- PDF-filerna exporteras till samma plats som arbetsboken är sparad i.
- Om det redan finns en fil med samma namn som det angivna filnamnet kommer den att ersättas.
- Makrot antar att kalkylbladen "sign.alla" och "sign.behandlad" finns i arbetsboken.
