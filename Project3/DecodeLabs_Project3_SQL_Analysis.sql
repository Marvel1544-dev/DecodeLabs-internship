-- DecodeLabs Data Analytics Internship
-- Project 3: SQL Data Analysis
-- Database: DecodeLabs_Project3

USE DecodeLabs_Project3;

SELECT *
FROM dbo.Sheet1$;

--1: SELECT
SELECT
    OrderID,
    Date,
    Product,
    Quantity,
    TotalPrice
FROM dbo.Sheet1$;

--2: WHERE
SELECT
    OrderID,
    Date,
    Product,
    Quantity,
    TotalPrice
FROM dbo.Sheet1$
WHERE Product = 'Laptop';

--3: ORDER BY
SELECT
    OrderID,
    Product,
    Quantity,
    TotalPrice
FROM dbo.Sheet1$
ORDER BY TotalPrice DESC;

--4: GROUP BY
SELECT
    Product,
    COUNT(*) AS OrderCount
FROM dbo.Sheet1$
GROUP BY Product
ORDER BY OrderCount DESC;

--5: SUM
SELECT
    Product,
    SUM(TotalPrice) AS TotalSales
FROM dbo.Sheet1$
GROUP BY Product
ORDER BY TotalSales DESC;

--6: AVG
SELECT
    Product,
    AVG(TotalPrice) AS AverageSales
FROM dbo.Sheet1$
GROUP BY Product
ORDER BY AverageSales DESC;

--7: HAVING
SELECT
    Product,
    COUNT(*) AS OrderCount,
    SUM(TotalPrice) AS TotalSales
FROM dbo.Sheet1$
GROUP BY Product
HAVING COUNT(*) > 170
ORDER BY OrderCount DESC;

--8: Percentage contribution of each product to total sales
SELECT
    Product,
    SUM(TotalPrice) AS TotalSales,
    ROUND(
        (SUM(TotalPrice) * 100.0) / (SELECT SUM(TotalPrice) FROM dbo.Sheet1$),
        2
    ) AS SalesPercentage
FROM dbo.Sheet1$
GROUP BY Product
ORDER BY SalesPercentage DESC;

--9: Orders by Status

SELECT
    OrderStatus,
    COUNT(*) AS OrderCount
FROM dbo.Sheet1$
GROUP BY OrderStatus
ORDER BY OrderCount DESC;

--10: Average order value by payment method

SELECT
    PaymentMethod,
    COUNT(*) AS OrderCount,
    AVG(TotalPrice) AS AverageOrderValue
FROM dbo.Sheet1$
GROUP BY PaymentMethod
ORDER BY AverageOrderValue DESC;

--11: Total sales by payment method
SELECT
    PaymentMethod,
    COUNT(*) AS OrderCount,
    SUM(TotalPrice) AS TotalSales
FROM dbo.Sheet1$
GROUP BY PaymentMethod
ORDER BY TotalSales DESC;

--12: Sales by Referral Source

SELECT
    ReferralSource,
    COUNT(*) AS OrderCount,
    SUM(TotalPrice) AS TotalSales,
    AVG(TotalPrice) AS AverageOrderValue
FROM dbo.Sheet1$
GROUP BY ReferralSource
ORDER BY TotalSales DESC;