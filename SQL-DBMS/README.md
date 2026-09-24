# SQL-DBMS: Rent a car database

Database project from the Database Management Systems course (SQL Server). The database stores cars, customers, cities, cantons and banks for a small rent a car company.

![ER diagram](ER%20dijagram.PNG.png)

## Scripts

Run them in this order in SQL Server Management Studio:

| Step | File | What it does |
|------|------|--------------|
| 1 | [01_create_database.sql](01_create_database.sql) | Creates the `RENTACAR` database (change the file paths to your machine) |
| 2 | [02_create_tbl_korisnik.sql](02_create_tbl_korisnik.sql) | Customers table |
| 3 | [03_create_tbl_grad.sql](03_create_tbl_grad.sql) | Cities table |
| 4 | [04_create_tbl_kanton.sql](04_create_tbl_kanton.sql) | Cantons table |
| 5 | [05_create_tbl_banka.sql](05_create_tbl_banka.sql) | Banks table |
| 6 | [06_create_tbl_auta.sql](06_create_tbl_auta.sql) | Cars table |
| 7 to 11 | `07_insert_*.sql` to `11_insert_auta.sql` | Sample data |
| 12 | [12_select_kantoni.sql](12_select_kantoni.sql) | List all cantons |
| 13 | [13_select_vrsta_goriva.sql](13_select_vrsta_goriva.sql) | Fuel type of every car |
| 14 | [14_select_bmw.sql](14_select_bmw.sql) | Search cars by brand with `LIKE` |

## Other files

- `RENTA A CAR DOKUMENT.docx`: project documentation
- `Kreiranje baze.png`, `UPIT1-3.png`: screenshots of the database and query results
- `iznajmiauto.mdf`: the SQL Server database file

## Cleanup notes

- Scripts were renamed from `SQLQuery1.sql` to `SQLQuery20.sql` to numbered, descriptive names in the order they should run.
- Removed two outdated drafts: an earlier version of the cars table (`SQLQuery2`) and a duplicate cantons table without the `dbo.TBL` prefix (`SQLQuery5`).
- Removed the transaction log file (`.LDF`), which SQL Server recreates.
- Removed about 200 KB of empty padding from the create database script.

A redesigned version with foreign keys, a rentals table, views and stored procedures is in [practice/sql](../practice/sql).
