-- Exercise 3: Queries on RENTACAR_V2
-- Topics: JOIN, LEFT JOIN, GROUP BY / HAVING, subqueries, date functions,
--         views, stored procedures and a transaction

USE RENTACAR_V2;
GO

-- 1. All cars with the city and canton where they are located
SELECT c.Brand, c.Model, c.PlateNumber, ci.Name AS City, ca.Name AS Canton
FROM dbo.Car c
JOIN dbo.City ci   ON ci.CityID = c.CityID
JOIN dbo.Canton ca ON ca.CantonID = ci.CantonID
ORDER BY ca.Name, ci.Name;

-- 2. Diesel cars cheaper than 65 KM per day
SELECT Brand, Model, PricePerDay
FROM dbo.Car
WHERE FuelType = N'DIZEL' AND PricePerDay < 65
ORDER BY PricePerDay;

-- 3. Customers without a bank (LEFT JOIN + IS NULL)
SELECT cu.FirstName, cu.LastName
FROM dbo.Customer cu
LEFT JOIN dbo.Bank b ON b.BankID = cu.BankID
WHERE b.BankID IS NULL;

-- 4. Number of rentals and total income per car (finished rentals only)
SELECT c.Brand, c.Model,
       COUNT(r.RentalID) AS Rentals,
       ISNULL(SUM((DATEDIFF(DAY, r.StartDate, r.EndDate) + 1) * c.PricePerDay), 0) AS Income
FROM dbo.Car c
LEFT JOIN dbo.Rental r ON r.CarID = c.CarID AND r.EndDate IS NOT NULL
GROUP BY c.Brand, c.Model
ORDER BY Income DESC;

-- 5. Customers with more than one rental (HAVING)
SELECT cu.FirstName, cu.LastName, COUNT(*) AS Rentals
FROM dbo.Customer cu
JOIN dbo.Rental r ON r.CustomerID = cu.CustomerID
GROUP BY cu.FirstName, cu.LastName
HAVING COUNT(*) > 1;

-- 6. Cars that were never rented (subquery with NOT EXISTS)
SELECT Brand, Model
FROM dbo.Car c
WHERE NOT EXISTS (SELECT 1 FROM dbo.Rental r WHERE r.CarID = c.CarID);

-- 7. Cars more expensive than the average price (subquery)
SELECT Brand, Model, PricePerDay
FROM dbo.Car
WHERE PricePerDay > (SELECT AVG(PricePerDay) FROM dbo.Car);
GO

-- 8. View: cars that are currently available
CREATE VIEW dbo.vAvailableCars AS
SELECT c.CarID, c.Brand, c.Model, c.PricePerDay, ci.Name AS City
FROM dbo.Car c
JOIN dbo.City ci ON ci.CityID = c.CityID
WHERE NOT EXISTS (
    SELECT 1 FROM dbo.Rental r
    WHERE r.CarID = c.CarID AND r.EndDate IS NULL
);
GO

SELECT * FROM dbo.vAvailableCars ORDER BY PricePerDay;
GO

-- 9. Stored procedure: rent a car (only if it is available)
CREATE PROCEDURE dbo.RentCar
    @CarID INT,
    @CustomerID INT,
    @StartDate DATE
AS
BEGIN
    SET NOCOUNT ON;

    IF NOT EXISTS (SELECT 1 FROM dbo.vAvailableCars WHERE CarID = @CarID)
    BEGIN
        RAISERROR(N'This car is not available.', 16, 1);
        RETURN;
    END

    INSERT INTO dbo.Rental (CarID, CustomerID, StartDate)
    VALUES (@CarID, @CustomerID, @StartDate);
END;
GO

-- 10. Stored procedure: return a car and calculate the price
CREATE PROCEDURE dbo.ReturnCar
    @RentalID INT,
    @EndDate DATE
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRANSACTION;

    UPDATE dbo.Rental
    SET EndDate = @EndDate
    WHERE RentalID = @RentalID AND EndDate IS NULL;

    IF @@ROWCOUNT = 0
    BEGIN
        ROLLBACK TRANSACTION;
        RAISERROR(N'Rental not found or already returned.', 16, 1);
        RETURN;
    END

    COMMIT TRANSACTION;

    SELECT r.RentalID,
           DATEDIFF(DAY, r.StartDate, r.EndDate) + 1 AS Days,
           (DATEDIFF(DAY, r.StartDate, r.EndDate) + 1) * c.PricePerDay AS TotalPrice
    FROM dbo.Rental r
    JOIN dbo.Car c ON c.CarID = r.CarID
    WHERE r.RentalID = @RentalID;
END;
GO

-- Try the procedures
EXEC dbo.RentCar @CarID = 2, @CustomerID = 3, @StartDate = '2024-06-10';
EXEC dbo.RentCar @CarID = 5, @CustomerID = 1, @StartDate = '2024-06-10';   -- error: already rented
EXEC dbo.ReturnCar @RentalID = 7, @EndDate = '2024-06-05';
