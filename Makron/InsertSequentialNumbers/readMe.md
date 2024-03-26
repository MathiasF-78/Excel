## InsertSequentialNumbers Macro

### Beskrivning
Detta makro används för att fylla i en kolumn med sekventiella nummer i stigande ordning i Excel.

### Användning
1. Markera den cell där du vill börja fylla i sekventiella nummer.
2. Kör makrot "InsertSequentialNumbers".

### Åtgärder som utförs av makrot:
- Makrot fyller den markerade cellen med värdet 1.
- Det utvidgar sedan det markerade området nedåt så att sekventiella nummer fylls i.
- Numreringen fortsätter i stigande ordning från det valda värdet, med ökning med 1 för varje rad nedåt.

### Anmärkningar
- Makrot antar att den markerade cellen är den första cellen där sekventiella nummer ska fyllas i.
- Numreringen fortsätter så länge det finns tomma celler under den första cellen.
- Om det finns data eller tomma celler längre ned kommer numreringen att avbrytas vid den punkten.

