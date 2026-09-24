-- Exercise 1: Rent a car database, version 2 (SQL Server / T-SQL)
--
-- A redesign of the original RENTACAR coursework database (see ../../SQL-DBMS).
-- Changes compared to the original:
--   * foreign keys between tables (the original had none)
--   * a separate RENTAL table, so one car can be rented many times
--   * CHECK and UNIQUE constraints
--   * a city belongs to a canton instead of storing both on the car

CREATE DATABASE RENTACAR_V2;
GO
USE RENTACAR_V2;
GO

CREATE TABLE dbo.Canton (
    CantonID INT IDENTITY PRIMARY KEY,
    Name     NVARCHAR(50) NOT NULL UNIQUE
);

CREATE TABLE dbo.City (
    CityID   INT IDENTITY PRIMARY KEY,
    Name     NVARCHAR(50) NOT NULL,
    CantonID INT NOT NULL REFERENCES dbo.Canton(CantonID)
);

CREATE TABLE dbo.Bank (
    BankID INT IDENTITY PRIMARY KEY,
    Name   NVARCHAR(50) NOT NULL UNIQUE
);

CREATE TABLE dbo.Customer (
    CustomerID INT IDENTITY PRIMARY KEY,
    FirstName  NVARCHAR(30) NOT NULL,
    LastName   NVARCHAR(30) NOT NULL,
    Phone      VARCHAR(20)  NULL,
    CityID     INT NOT NULL REFERENCES dbo.City(CityID),
    BankID     INT NULL     REFERENCES dbo.Bank(BankID)
);

CREATE TABLE dbo.Car (
    CarID        INT IDENTITY PRIMARY KEY,
    Brand        NVARCHAR(30) NOT NULL,
    Model        NVARCHAR(30) NOT NULL,
    ModelYear    SMALLINT     NOT NULL CHECK (ModelYear BETWEEN 1990 AND 2100),
    FuelType     NVARCHAR(10) NOT NULL CHECK (FuelType IN (N'DIZEL', N'BENZIN', N'HIBRID', N'STRUJA')),
    PlateNumber  VARCHAR(15)  NOT NULL UNIQUE,
    PricePerDay  DECIMAL(8,2) NOT NULL CHECK (PricePerDay > 0),
    CityID       INT NOT NULL REFERENCES dbo.City(CityID)
);

CREATE TABLE dbo.Rental (
    RentalID   INT IDENTITY PRIMARY KEY,
    CarID      INT  NOT NULL REFERENCES dbo.Car(CarID),
    CustomerID INT  NOT NULL REFERENCES dbo.Customer(CustomerID),
    StartDate  DATE NOT NULL,
    EndDate    DATE NULL,             -- NULL = car is still rented
    CONSTRAINT CK_Rental_Dates CHECK (EndDate IS NULL OR EndDate >= StartDate)
);
GO
