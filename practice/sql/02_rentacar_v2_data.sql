-- Exercise 2: Sample data for RENTACAR_V2
-- All names and phone numbers are made up.

USE RENTACAR_V2;
GO

INSERT INTO dbo.Canton (Name) VALUES
(N'SBK'), (N'Sarajevski'), (N'HNK'), (N'ZDK'), (N'Kanton 10');

INSERT INTO dbo.City (Name, CantonID) VALUES
(N'Travnik', 1), (N'Vitez', 1), (N'Sarajevo', 2), (N'Mostar', 3), (N'Zenica', 4), (N'Livno', 5);

INSERT INTO dbo.Bank (Name) VALUES
(N'Sparkasse'), (N'Raiffeisen'), (N'BBI'), (N'UniCredit');

INSERT INTO dbo.Customer (FirstName, LastName, Phone, CityID, BankID) VALUES
(N'Amir',  N'Hadzic',   '061-000-001', 1, 1),
(N'Lejla', N'Kovac',    '061-000-002', 3, 2),
(N'Tarik', N'Begic',    '061-000-003', 4, NULL),
(N'Selma', N'Mehic',    '061-000-004', 5, 3),
(N'Kenan', N'Softic',   '061-000-005', 1, 4);

INSERT INTO dbo.Car (Brand, Model, ModelYear, FuelType, PlateNumber, PricePerDay, CityID) VALUES
(N'Audi',       N'A4',      2012, N'DIZEL',  'A01-K-001', 60.00, 1),
(N'Volkswagen', N'Passat',  2017, N'BENZIN', 'A01-K-002', 55.00, 3),
(N'Maserati',   N'Ghibli',  2020, N'BENZIN', 'A01-K-003', 250.00, 3),
(N'Skoda',      N'Octavia', 2019, N'DIZEL',  'A01-K-004', 45.00, 4),
(N'Toyota',     N'Yaris',   2022, N'HIBRID', 'A01-K-005', 40.00, 5),
(N'BMW',        N'320d',    2018, N'DIZEL',  'A01-K-006', 70.00, 1),
(N'Renault',    N'Clio',    2015, N'BENZIN', 'A01-K-007', 35.00, 6);

INSERT INTO dbo.Rental (CarID, CustomerID, StartDate, EndDate) VALUES
(1, 1, '2024-03-01', '2024-03-05'),
(2, 2, '2024-03-10', '2024-03-12'),
(4, 3, '2024-04-01', '2024-04-08'),
(1, 4, '2024-04-15', '2024-04-16'),
(6, 1, '2024-05-02', '2024-05-09'),
(3, 5, '2024-05-20', '2024-05-21'),
(5, 2, '2024-06-01', NULL);          -- still rented
GO
