# Visual Basic 6

Desktop applications from the Visual Basic 6 course. Open the `.vbp` project file in the VB6 IDE to run a project.

| Project | Description |
|---------|-------------|
| [Bankomat](Bankomat) | ATM simulator with customers, transactions and a statement report, backed by an Access database |
| [Balkanske valute konvertor](Balkanske%20valute%20konvertor) | Currency converter for Balkan currencies |
| [Baza podataka UI--](Baza%20podataka%20UI--) | Contacts database with a simple form UI (Access database) |
| [Enkripcija-dekripcija podataka](Enkripcija-dekripcija%20podataka) | Text encryption and decryption tool |
| [Glitchanje misa](Glitchanje%20misa) | "Dancing mouse" prank app that moves the cursor around |
| [Najveći i najmanji broj](Najve%C4%87i%20i%20najmanji%20broj) | Find the largest and smallest number |
| [Paint u vb](Paint%20u%20vb) | Paint clone with brushes, colours and custom cursors |
| [Prijava i odjava ispita---](Prijava%20i%20odjava%20ispita---) | Exam registration and deregistration form |
| [Puzzle igra](Puzzle%20igra) | Classic sliding puzzle game with sounds and options |
| [Web browser V2](Web%20browser%20V2) | Simple web browser built on the WebBrowser control |

`Prezentacija.pptx` is the course presentation.

## Cleanup notes

- Removed files that don't belong in source control: a compiled `.exe`, IDE workspace files (`.vbw`), source control metadata (`MSSCCPRJ.SCC`), log files, a temp file and `Thumbs.db`.
- Three files (`Web browser V2/Form1.frm`, `Prijava i odjava ispita---/Form1.frm`, `Paint u vb/Module1.bas`) were padded with hundreds of KB of empty bytes. The padding was removed; the code is unchanged.
