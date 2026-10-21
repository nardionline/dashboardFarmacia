# Dergesa Reports

Aplikacion Windows Forms për raporte magazine nga databaza ekzistuese Microsoft Access 2003 (`.mdb`). Nuk ndryshon strukturën e tabelave.

## Raportet
- Hyrje dhe dalje sipas artikullit, me filtër datash.
- Gjendja aktuale nga `Artikujt.Gjendjemalli`.
- Kartela e lëvizjeve për një artikull të zgjedhur, me filtër datash.
- Eksport i rezultatit në CSV për Excel.

## Kërkimi i artikullit
Lista dhe plotësimi automatik kërkojnë vetëm sipas emërtimit `Artikulli`; Kodi dhe Barcode nuk përdoren si fusha kërkimi.

## Databaza
Aplikacioni lejon zgjedhjen e skedarit `.mdb` ose `.accdb` nga butoni **Zgjidh MDB...**. Parazgjedhja është `dergesa.mdb` në të njëjtën dosje me EXE-në.

Connection string:
`Provider=Microsoft.ACE.OLEDB.12.0;Data Source=<database path>;Persist Security Info=False;`

Build është **x86**. Në kompjuter duhet të jetë instaluar ACE OLEDB 32-bit. Kërkesat SQL shmangin `Nz()` dhe janë përshtatur për fushat e databazës Access të projektit.

## Ndërtimi përmes GitHub Actions
1. Ngarko të gjitha skedarët e këtij ZIP-i në root të repository-t.
2. Bëj commit/push në `main` ose `master`.
3. Hape skedën **Actions** dhe prit workflow `Build Dergesa Reports EXE`.
4. Shkarko artifact `DergesaReports-Windows-x86` nga run-i i përfunduar.
5. Ekstrakto ZIP-in dhe ekzekuto `DergesaReports.exe`.

## Shënime
- Filtri i datave përfshin të gjithë ditën e fundit të zgjedhur.
- Faturat e anuluara përjashtohen (`Anulluar=True`).
- Raporti i gjendjes lexon gjendjen aktuale të ruajtur në `Artikujt.Gjendjemalli`; nuk llogarit gjendjen historike në një datë të kaluar.
- Projekti duhet testuar kundrejt kopjes së databazës reale pas build-it.
