# VBA Project Unprotect Hook

Detta projekt innehåller VBA-kod som demonstrerar en avancerad teknik för att **detektera och manipulera skyddet på ett VBA-projekt** i Excel via Windows API:er och låg-nivå minneshantering.

> ⚠️ **Observera:** Den här koden är för **utbildningssyfte** och **ska inte användas i produktionsmiljö eller för att kringgå säkerhetsskydd i tredje parts-filer.** Det bryter mot god etik och eventuellt gällande licensavtal.

---

## 🧩 Innehåll

### Module1: `Hook`-funktion och minnespatchning

Denna modul använder sig av följande Windows API-funktioner:
- `VirtualProtect`
- `MoveMemory`
- `GetProcAddress`
- `DialogBoxParam`

Syftet är att:
- Hämta adressen till `DialogBoxParamA` från `user32.dll`
- Spara originalbytes (för återställning)
- Skriva in egna bytes (hook) för att fånga upp dialoger som används när VBA-projektet skyddas
- Återställa originalkoden efter manipulation

### Module2: `unprotected`-rutin

En enkel procedur som anropar `Hook` och visar en `MsgBox` om VBA-projektet är oskyddat:

```vba
If Hook Then
    MsgBox "VBA Project is unprotected!", vbInformation, "*****"
End If
