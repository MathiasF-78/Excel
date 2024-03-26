# AddPrefix Makro

Detta VBA-makro för Excel, `AddPrefix`, låter användare lägga till en textsträng som ett prefix till valda cellers innehåll. Detta kan vara särskilt användbart när du behöver markera eller kategorisera data på ett enhetligt sätt, till exempel genom att lägga till en kod eller en kort beskrivning framför befintlig text.

## Funktioner

- **Lägg till Prefix:** Möjliggör för användare att dynamiskt lägga till ett valfritt prefix till början av varje vald cell i ett Excel-ark.
- **Användarinput för Prefix:** Användaren uppmanas att mata in den text som ska läggas till som prefix, vilket ger flexibilitet för olika användningsområden.

## Användningsinstruktioner

1. **Markera Cellerna:** Börja med att markera de celler i ditt Excel-ark där du vill lägga till ett prefix.
2. **Kör Makrot:** Kör `AddPrefix`-makrot. Detta kan du göra genom att gå till `Utvecklare`-fliken i Excel, välja `Makron`, leta upp `AddPrefix`, och sedan klicka på `Kör`.
3. **Ange Prefix:** När makrot körs kommer en dialogruta att visas där du ombeds att ange det prefix du vill lägga till. Skriv in det önskade prefixet och klicka på OK.
4. **Resultat:** Prefixet du angav läggs nu till i början av varje cell du markerade.

## Noteringar

- Makrot ignorerar celler som är tomma och lägger endast till prefixet i celler som redan har innehåll.
- Detta makro använder inte ett mellanslag mellan prefixet och cellinnehållet. Om du behöver ett mellanslag mellan ditt prefix och den ursprungliga texten, kan du enkelt lägga till det i dialogrutan för prefixinput.

## Anpassningar

Om du vill modifiera makrot för att inkludera ett mellanslag mellan prefixet och cellinnehållet, kan du redigera raden:

```vb
Rng.Value = addStr & Rng.Value
```

till:

```vb
Rng.Value = addStr & " " & Rng.Value
```

Genom att följa dessa steg, kan du effektivt lägga till prefix i ditt Excel-ark med hjälp av `AddPrefix`-makrot.
```

Denna README.md ger en översikt av makrots funktion, dess användning, och tips för anpassning. Modifiera innehållet efter behov baserat på makrots specifika beteende och dina önskemål.
