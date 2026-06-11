# PDF_To_Word_Converter

**PDF til Word – helt lokalt, uten opplasting.**

Dra en PDF inn i vinduet, og Word-filen lagres automatisk i samme mappe. Ingen opplasting, ingen sky, ingen konto. Alt skjer på din egen maskin – også tekstgjenkjenning av skannede dokumenter.

## For brukere

1. Last ned **PDF_To_Word_Converter.exe** og legg den der du vil (Skrivebordet fungerer fint).
2. Dobbeltklikk for å starte. Ingen installasjon.
3. Dra en PDF inn i vinduet – eller klikk for å velge fil.
4. Word-filen dukker opp i samme mappe som PDF-en, med samme navn.

**Første gang:** Windows SmartScreen kan vise «Windows beskyttet PC-en din». Klikk **Mer info** → **Kjør likevel**. Dette skjer fordi programmet ikke er kodesignert, ikke fordi noe er galt.

## Hva programmet gjør

| Dokumenttype | Behandling |
|---|---|
| Digital PDF (laget fra Word, nettside o.l.) | Direkte konvertering med bevart layout, tabeller, bilder og formatering |
| Skannet PDF (papir som er skannet) | Oppdages automatisk. Tekstgjenkjenning (OCR) på norsk og engelsk, deretter gjenoppbygging med avsnitt og overskrifter |

Finnes Word-filen fra før, lagres den nye som «navn (2).docx» – ingenting overskrives.

## Personvern

PDF_To_Word_Converter trenger ikke internett og sender ingenting ut av maskinen. Dokumentene behandles i minnet og i en midlertidig mappe som slettes etterpå. Det gjør programmet trygt også for dokumenter med personopplysninger.

## Begrensninger

- Skannede dokumenter gjenskapes som tekst med avsnitt og overskrifter – ikke pikselidentisk kopi av originalen.
- Svært komplekse layouter (flerspaltede sider, roterte tabeller) kan kreve litt etterjustering i Word.
- Passordbeskyttede PDF-er må låses opp først.
- OCR-kvaliteten følger skannekvaliteten: skjeve, mørke eller uskarpe skanninger gir dårligere resultat.

## For deg som skal bygge eller videreutvikle

Prosjektet er ren Python (Tkinter + pdf2docx + PyMuPDF + Tesseract).

- **Bygge ferdig .exe uten Windows-maskin:** se [GUIDE-GITHUB-ACTIONS.md](GUIDE-GITHUB-ACTIONS.md)
- **Bygge lokalt på Windows:** installer Python 3.10+ og Tesseract (`choco install tesseract`), kjør `build-windows.bat`. Ferdig fil havner i `dist\PDF_To_Word_Converter.exe`.
- **Kjøre fra kildekode (utvikling):** `pip install -r requirements.txt`, deretter `python app.py`. For OCR i utvikling må Tesseract ligge i PATH.

### Filoversikt

```
app.py                       Grensesnittet (dra-og-slipp-vindu)
converter.py                 Konverteringslogikk og OCR
requirements.txt             Python-avhengigheter
build-windows.bat            Lokal bygging på Windows
.github/workflows/bygg.yml   Automatisk bygging via GitHub Actions
GUIDE-GITHUB-ACTIONS.md      Steg-for-steg-guide for bygging uten Windows-maskin
```

## Innstillinger

**Fjern topp- og bunntekst** (på som standard): fjerner automatisk linjer som gjentas øverst/nederst på sidene (dokumentnavn, «Konfidensielt» o.l.) og sidetall, slik at de ikke havner midt i brødteksten i Word. Skru av med avkrysningsboksen hvis du vil beholde dem.

## Store dokumenter

Konverteringen kjører i en egen prosess, så programmet fryser ikke og overlever selv om et svært stort eller komplekst dokument feiler – du får en feilmelding og kan avbryte underveis med Avbryt-knappen. Svært store dokumenter (hundrevis av sider, mye grafikk) kan likevel ta flere minutter; del dem gjerne opp.
