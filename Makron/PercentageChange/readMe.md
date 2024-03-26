## Percent_Change_Formula Macro

### Beskrivning
Detta makro skapar en formel för att beräkna procentuell förändring mellan två värden. Formeln jämför det nya värdet med det gamla värdet och beräknar sedan procentuell förändring.

### Användning
1. Kör makrot "Percent_Change_Formula".
2. Följ anvisningarna i de dialogrutor som visas för att markera det nya och det gamla värdet.
3. Makrot skapar en formel i den aktiva cellen för att beräkna procentuell förändring mellan de två värdena.

### Anmärkningar
- Makrot förutsätter att användaren markerar det nya och det gamla värdet i två separata dialogrutor som visas när makrot körs.
- Formeln som skapas jämför det nya värdet med det gamla värdet och beräknar sedan procentuell förändring enligt följande: ((nytt värde - gammalt värde) / gammalt värde).
- Om det inte går att beräkna procentuell förändring på grund av att det gamla värdet är noll, kommer formeln att returnera 0.
- Formeln formateras automatiskt som procentuell förändring med två decimaler och visas som ett procenttecken.
