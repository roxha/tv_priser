# TV-priser – automatisk kjøring med GitHub Actions

Denne pakken bygger på v5.1: Elkjøp, Power og Elon. Komplett er deaktivert på grunn av dokumenterte nettverksfeil. Priser er ikke live-testet fra GitHub Actions; butikker kan blokkere skynettverk.

## Oppsett (ca. 10 minutter)

1. Logg inn på github.com og opprett et **privat** repository, f.eks. `tv-priser`. Velg **ikke** «Add README» ved opprettelsen.
2. Pakk ut ZIP-filen. Last opp **alle** filene og mappene til repositoryet, inkludert `.github/workflows/tv-priser.yml` (mappen `.github` er skjult i noen filvisninger). GitHub-nettsiden lar deg dra inn filer og mapper.
3. For å beholde gammel historikk: Kopier eksisterende `prisobservasjoner.csv` fra `C:\Program Files\Tv priser\` inn i `data/` i repositoryet, **før første kjøring**. Det er denne filen scriptet bygger videre på. `TV_priser.xlsx` genereres på nytt.
4. Gå til **Settings → Actions → General → Workflow permissions**. Velg **Read and write permissions** og lagre. Hvis dette valget er låst av organisasjonen, bruk et privat repository der du har kontroll.
5. Gå til **Actions → TV-priser daglig → Run workflow** for å prøve en manuell kjøring.
6. Se resultatet i **Actions**: grønn kjøring er vellykket; åpne kjøringen for logger. Under **Artifacts** ligger `TV-priser-rapport` med Excel, CSV og logg. Rapporten blir også oppdatert i `data/` via automatisk commit.
7. Planlagt kjøring er **07:00 UTC hver dag**, dvs. kl. **08:00 vintertid / 09:00 sommertid i Norge**. GitHub kan forsinke planlagte kjøringer; cron er ikke en garanti for eksakt starttid.

## Viktig

- Gratis GitHub Actions-kapasitet og grenser avhenger av kontotype og gjeldende GitHub-vilkår. Kontroller kvoten på GitHub.
- Repositoryet bør være privat: det inneholder historikk og kan inneholde tilpassede produktlenker.
- Hvis du bruker `produktlenker.txt`, legg den i `data/`.
- GitHub kan deaktivere planlagte workflows i inaktive repositoryer etter lengre tid. Se status i Actions ved behov.
- En vellykket jobb betyr ikke at alle priser ble funnet. Sjekk `data/innhenting.log` og fanen «Kjøringsstatus».
- Dersom du ønsker kjøring **08:00 norsk tid også om sommeren**, må cron-tidspunktet endres ved sommertid, eller man må bruke en scheduler med tidssonefunksjon.
- Bruk av sky-IP kan gi andre resultater enn fra Windows-PC-en din.
