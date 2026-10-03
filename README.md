# datasette-xlsx

Plugin till [Datasette](https://datasette.io/) som exporterar tabeller, vyer
och frågeresultat som Excel-filer (`.xlsx`).

Pluginet registrerar filändelsen `.xlsx` på samma sätt som Datasettes inbyggda
`.json`. Därför dyker en länk till Excel-exporten upp automatiskt på tabell-,
vy-, fråge- och radsidor.

## Installation

Pluginet finns inte på PyPI. Installera det från GitHub i samma miljö som
Datasette körs i:

```bash
datasette install git+https://github.com/svkau/datasette-xlsx
```

eller, i ett projekt som hanteras med uv:

```bash
uv add git+https://github.com/svkau/datasette-xlsx
```

Kontrollera att pluginet är laddat med `datasette plugins`.

## Användning

Lägg till `.xlsx` på en sidas adress, eller klicka på länken `xlsx` på sidan:

| Sida | Exempel |
|---|---|
| Tabell eller vy | `/databas/tabell.xlsx` |
| Filtrerad och sorterad tabell | `/databas/tabell.xlsx?år__gte=1950&_sort=namn` |
| SQL-fråga | `/databas.xlsx?sql=select+...` |
| Sparad fråga | `/databas/fraga.xlsx` |
| Enskild rad | `/databas/tabell/17.xlsx` |

## Vad exporten innehåller

**Hela resultatet, inte bara första sidan.** För tabeller och vyer följer
pluginet Datasettes egen sidindelning, så all filtrering, sökning, sortering
och kolumnval görs av Datasette och blir samma som på webbsidan. SQL-frågor och
sparade frågor körs om direkt mot databasen (skrivskyddat), så att resultatet
inte kapas vid `max_returned_rows`.

**Behörigheter följer med.** Anroparens cookies och `Authorization`-header
skickas vidare, så den som inte får se en tabell i Datasette kan inte heller
exportera den. Sparade frågor som skriver till databasen kan inte exporteras.

**Formatering:**

- Datat i varje blad formateras som en Excel-tabell (`tbl_<namn>`) med
  filterknappar och randade rader. Rubrikraden är låst vid scrollning.
- Kolumnbredderna anpassas efter rubrikerna och de första 1000 raderna, med
  bredd mellan 6 och 60 tecken.
- Heltal med 12–15 siffror, till exempel personnummer, visas fullt ut i stället
  för i exponentform (`1,99001E+11`) och går fortfarande att räkna med.

**Bladet "Om uttaget"** beskriver exporten: databas, tabell eller fråga,
SQL, adress, användare, tidpunkt, Datasette-version och antal rader. Där står
också om exporten har kapats.

## Begränsningar och ändrade värden

Excel har gränser som exporten måste anpassas till. Det här är de fall där
filen inte är en exakt kopia av datat i Datasette:

| Situation | Hantering |
|---|---|
| Fler än 1 048 575 rader | Fortsätter på ett nytt blad, med en egen tabell |
| Fler rader än `max_rows` | Exporten stoppas vid gränsen; noteras i "Om uttaget" |
| Text längre än 32 767 tecken | Kapas och markeras `…[TRUNKERAT]`; antal anges i "Om uttaget" |
| Heltal med 16 siffror eller fler | Sparas som text, eftersom Excel inte kan lagra dem exakt |
| Binärdata (BLOB) | Ersätts med `<binärdata, N byte>` |
| Tecken som inte är tillåtna i xlsx (styrtecken) | Tas bort |
| Text som börjar med `=` | Sparas som text, aldrig som formel |
| Kolumnnamn som är tomma eller förekommer flera gånger | Döps om, t.ex. `id`, `id` → `id`, `id_2`, och tom → `Kolumn3` |
| Radbrytning i kolumnnamn | Ersätts med mellanslag |

Inledande nollor (`0701234567`) finns bara kvar om värdet är lagrat som text i
databasen. Är det lagrat som heltal har nollan redan försvunnit i SQLite.

## Konfiguration

Inställningarna läggs under `plugins` i Datasettes metadata
(`metadata.yaml`/`metadata.json`, eller `datasette.yaml` i Datasette 1.0).
De kan anges för hela instansen, per databas eller per tabell.

```yaml
plugins:
  datasette-xlsx:
    max_rows: 1000000     # högsta antal rader i en export (standard 1 000 000)
    time_limit_ms: 30000  # tidsgräns för SQL-frågor i ms (standard 30 000)
```

## Att tänka på

- **Belastning.** SQL-frågor körs om utan Datasettes radgräns. Den som får
  köra egen SQL kan alltså exportera upp till `max_rows` rader, med tidsgränsen
  `time_limit_ms`. Sänk värdena om instansen är publik.
- **Personuppgifter.** Användarens id (`actor.id`) skrivs in i bladet
  "Om uttaget". Exporterade filer lämnar Datasettes behörighetskontroll och kan
  spridas vidare. Tänk på det när databasen innehåller personuppgifter.

## Utveckling

```bash
uv sync --extra test
uv run pytest
```

Testa manuellt mot en lokal databas, som bör ligga utanför repot:

```bash
uv run datasette ../datasette-testdata/exempel.db
```

## Licens

MIT, se [LICENSE](LICENSE).
