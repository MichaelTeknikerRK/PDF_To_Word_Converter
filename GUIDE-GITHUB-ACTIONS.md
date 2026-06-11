# Guide: Bygg PDF_To_Word_Converter.exe med GitHub Actions

Denne guiden viser hvordan du får en ferdig **PDF_To_Word_Converter.exe** uten å bygge noe selv. GitHub gjør jobben på sine egne Windows-maskiner – gratis, og du trenger bare en nettleser.

Tidsbruk: ca. 10 minutter første gang, 0 minutter senere (bygging skjer automatisk ved hver endring).

## Steg 1: Opprett et repo

1. Logg inn på [github.com](https://github.com) og klikk **+** øverst til høyre → **New repository**.
2. Navn: `pdf-to-word-converter` (eller hva du vil).
3. Velg **Private** hvis koden ikke skal være offentlig. Actions er gratis også for private repoer (2 000 minutter/måned – ett bygg tar ca. 5).
4. Klikk **Create repository**.

## Steg 2: Last opp filene

**Enkleste vei (nettleser):**

1. I det nye repoet: klikk **uploading an existing file** (lenken på forsiden).
2. Dra inn alle filene fra prosjektmappen: `app.py`, `converter.py`, `requirements.txt`, `build-windows.bat`, `README.md`, `GUIDE-GITHUB-ACTIONS.md`.
3. Klikk **Commit changes**.

Mappen `.github/workflows/` kan ikke dras inn som mappe i nettleseren, så den lages manuelt:

4. Klikk **Add file** → **Create new file**.
5. I navnefeltet skriver du nøyaktig: `.github/workflows/bygg.yml`
   (GitHub lager mappene automatisk når du skriver skråstrekene.)
6. Lim inn innholdet fra `bygg.yml` i prosjektmappen.
7. Klikk **Commit changes**.

**Alternativ (kommandolinje), hvis du har git installert:**

```bash
cd pdf2word
git init && git add . && git commit -m "Første versjon av PDF_To_Word_Converter"
git branch -M main
git remote add origin https://github.com/DITT-BRUKERNAVN/pdf-to-word-converter.git
git push -u origin main
```

## Steg 3: Vent på bygget

1. Klikk på fanen **Actions** øverst i repoet.
2. Du ser kjøringen **«Bygg PDF_To_Word_Converter (Windows-exe)»** med en gul prikk (pågår). Den startet automatisk da du lastet opp.
3. Etter 4–6 minutter blir prikken en grønn hake. Klikk på kjøringen.

> Hvis Actions er deaktivert i repoet, får du en knapp om å aktivere det første gang – klikk **I understand my workflows, go ahead and enable them**.

## Steg 4: Last ned PDF_To_Word_Converter.exe

1. Inne på den ferdige kjøringen: bla ned til seksjonen **Artifacts**.
2. Klikk på **PDF_To_Word_Converter-Windows** – en zip-fil lastes ned.
3. Pakk ut zip-en. Inni ligger **PDF_To_Word_Converter.exe** – klar til å deles med kolleger.

Filen er rundt 80–120 MB fordi tekstgjenkjenning (Tesseract med norsk og engelsk) er bakt inn. Det er prisen for at brukerne slipper all installasjon.

## Senere endringer

Hver gang du endrer en fil i repoet (f.eks. redigerer `app.py` direkte i nettleseren og committer), bygges en ny .exe automatisk. Du kan også starte et bygg manuelt: **Actions** → velg workflowen → **Run workflow**.

## Feilsøking

| Problem | Løsning |
|---|---|
| Ingen kjøring vises under Actions | Sjekk at filen ligger nøyaktig på stien `.github/workflows/bygg.yml`, og at du committet til grenen `main` |
| Rød X (bygget feilet) | Klikk på kjøringen → klikk på steget med rød X for å se feilmeldingen. Vanligste årsak er skrivefeil i yml-filen ved manuell innliming |
| Artefakten er borte | Artefakter slettes automatisk etter 90 dager. Kjør workflowen på nytt (Run workflow) for en fersk .exe |
| SmartScreen-advarsel hos brukerne | Normalt for usignerte programmer: «Mer info» → «Kjør likevel». For bred utrulling i kommunen: be IT hviteliste filen, eller vurder kodesignering |

## Tips for distribusjon internt

Legg den ferdige `PDF_To_Word_Converter.exe` på et felles område (Teams/SharePoint/fellesdisk) sammen med én setning: *«Dra en PDF inn i vinduet – Word-filen havner i samme mappe.»* Mer opplæring trengs ikke.
