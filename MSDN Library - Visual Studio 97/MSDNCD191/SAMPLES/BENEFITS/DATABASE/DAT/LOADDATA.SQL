--Corporate Benefits Sample
--==========================
--by Ken Bergmann
--
--Load Sample Dependant Type Domain
Print 'INSERTING Dependant Type Records...'
GO
INSERT INTO tbDepType(Descr, Code) VALUES ('Spouse/Domestic Partner' , 'PARTNER')
GO
INSERT INTO tbDepType(Descr, Code) VALUES ('Child' , 'CHILD')
GO
INSERT INTO tbDepType(Descr, Code) VALUES ('Parent' , 'PARENT')
GO
INSERT INTO tbDepType(Descr, Code) VALUES ('Sibling' , 'SIBLING')
GO
INSERT INTO tbDepType(Descr, Code) VALUES ('Other Relative' , 'RELATIV')
GO
INSERT INTO tbDepType(Descr, Code) VALUES ('Friend' , 'FRIEND')
GO
INSERT INTO tbDepType(Descr, Code) VALUES ('Estate' , 'ESTATE')
GO
Print '------------------------------------------------>Completed.'
GO

--Load Sample Benefit Tax Type Domain
PRINT 'INSERTING Benefit Tax Records...'
GO
INSERT INTO tbBnfTax(Descr, Code) VALUES ('Not Applicable' , 'TAXNA')
GO
INSERT INTO tbBnfTax(Descr, Code) VALUES ('Pre-Tax' , 'TAXPRE')
GO
INSERT INTO tbBnfTax(Descr, Code) VALUES ('Post-Tax' , 'TAXPOST')
GO
INSERT INTO tbBnfTax(Descr, Code) VALUES ('Employee Choice' , 'TAXCHCE')
GO
Print '------------------------------------------------>Completed.'
GO

--Load Sample Benefit Type Domain
PRINT 'INSERTING Benefit Type Records...'
GO
INSERT INTO tbBnfType(Descr, Code) VALUES ('Medical' , 'MEDICAL')
GO
INSERT INTO tbBnfType(Descr, Code) VALUES ('Life Insurance' , 'LIFEINS')
GO
INSERT INTO tbBnfType(Descr, Code) VALUES ('Disability Insurance' , 'DISABL')
GO
INSERT INTO tbBnfType(Descr, Code) VALUES ('Retirement' , 'PENSION')
GO
INSERT INTO tbBnfType(Descr, Code) VALUES ('Perks' , 'PERKS')
GO
Print '------------------------------------------------>Completed.'
GO

--Load Sample Employee Benefit Tax Type Domain
PRINT 'INSERTING Employee Benefit Tax Records...'
GO
INSERT INTO tbEmplBnfTax(Descr, Code) VALUES ('Not Applicable' , 'TAXNA')
GO
INSERT INTO tbEmplBnfTax(Descr, Code) VALUES ('Pre-Tax' , 'TAXPRE')
GO
INSERT INTO tbEmplBnfTax(Descr, Code) VALUES ('Post-Tax' , 'TAXPOST')
GO
Print '------------------------------------------------>Completed.'
GO

--Load Sample Employee Benefit Coverage Domain
PRINT 'INSERTING Employee Benefit Coverage Records...'
GO
INSERT INTO tbEmplBnfCoverage(Descr, Code) VALUES ('Employee Only' , 'EMPONLY')
GO
INSERT INTO tbEmplBnfCoverage(Descr, Code) VALUES ('Employee & Spouse/Domestic Partner' , 'EMPSPSE')
GO
INSERT INTO tbEmplBnfCoverage(Descr, Code) VALUES ('Employee & Eligible Children' , 'EMPCHLD')
GO
INSERT INTO tbEmplBnfCoverage(Descr, Code) VALUES ('Employee, Spouse & Eligible Children' , 'EMPALL')
GO
Print '------------------------------------------------>Completed.'
GO

--Load Sample Bank Data
PRINT 'INSERTING Bank Records...'
GO
INSERT INTO tbBank(Address1, Address2, City, Code, Name) 
VALUES ('1316 Acacia Street' , NULL , 'Helden' , '23-98-12' , 'First National Bank')
GO
INSERT INTO tbBank(Address1, Address2, City, Code, Name) 
VALUES ('48 Bruning Street' , NULL , 'Leyton' , '23-89-43' , 'Highland Bank')
GO
INSERT INTO tbBank(Address1, Address2, City, Code, Name) 
VALUES ('435 Elgar Road' , 'PO Box 1000' , 'Winnersh' , '12-23-31' , 'Midshires Bank')
GO
INSERT INTO tbBank(Address1, Address2, City, Code, Name) 
VALUES ('233 Stewart Way' , NULL , 'Redmond' , '94-73-93' , 'SDS Credit Union')
GO
INSERT INTO tbBank(Address1, Address2, City, Code, Name) 
VALUES ('342 Heddon Ave' , 'Hill Valley' , 'KempTown' , '90-98-23' , 'Commerce Bank')
GO
INSERT INTO tbBank(Address1, Address2, City, Code, Name) 
VALUES ('22 King Blvd' , NULL , 'Cookley' , '03-81-72' , 'Tower Buiding Society')
GO
Print '------------------------------------------------>Completed.'
GO

--Load Sample Employee Data
PRINT 'INSERTING Employee Records...'
GO
INSERT INTO tbEmployee(BankID, Address1, Address2, BankAccount, City, DateOfBirth, EndDate, FirstName, Gender, LastName, Position, Salary, StartDate, UserName)  
VALUES (4 , '27 Washington Street' , NULL , '093829060333' , 'Redmond' , '7/4/55' , NULL , 'Willemina' , 1 , 'Smith' , 'Chief Executive Officer' , 55000 , '1/1/82' , 'bills')
GO
INSERT INTO tbEmployee(BankID, Address1, Address2, BankAccount, City, DateOfBirth, EndDate, FirstName, Gender, LastName, Position, Salary, StartDate, UserName)  
VALUES (5 , '48 Summer St.' , NULL , '0846-84522465' , 'KempTown' , '10/5/56' , NULL , 'Larry' , 0 , 'Fine' , 'Benefits Coordinator' , 37599 , '4/17/89' , 'larryf')
GO
INSERT INTO tbEmployee(BankID, Address1, Address2, BankAccount, City, DateOfBirth, EndDate, FirstName, Gender, LastName, Position, Salary, StartDate, UserName)  
VALUES (4 , '44 Madison Ave' , NULL , '1-87362586576' , 'Bellevue' , '6/19/57' , NULL , 'Moe' , 0 , 'Howard' , 'Product Planner' , 39766 , '8/4/91' , 'moeh')
GO
INSERT INTO tbEmployee(BankID, Address1, Address2, BankAccount, City, DateOfBirth, EndDate, FirstName, Gender, LastName, Position, Salary, StartDate, UserName)  
VALUES (2 , '100 S. Baker Drive' , NULL , '96-0576035' , 'Woodinville' , '3/17/60' , NULL , 'Shirley' , 1 , 'Harrison' , 'VP Finance' , 32801 , '5/27/84' , 'shirleyh')
GO
INSERT INTO tbEmployee(BankID, Address1, Address2, BankAccount, City, DateOfBirth, EndDate, FirstName, Gender, LastName, Position, Salary, StartDate, UserName)  
VALUES (6 , '66 Adams Blvd' , NULL , '3046235675' , 'Cookley' , '8/12/62' , NULL , 'Joe' , 0 , 'Besser' , 'Accountant' , 29333 , '12/27/87' , 'joeb')
GO
INSERT INTO tbEmployee(BankID, Address1, Address2, BankAccount, City, DateOfBirth, EndDate, FirstName, Gender, LastName, Position, Salary, StartDate, UserName)  
VALUES (5 , '89 Jefferson Way' , NULL , '23067237-00' , 'Duvalle' , '10/22/65' , NULL , 'Carol' , 1 , 'Simpson' , 'Production Manager' , 38034 , '1/3/84' , 'carols')
GO
INSERT INTO tbEmployee(BankID, Address1, Address2, BankAccount, City, DateOfBirth, EndDate, FirstName, Gender, LastName, Position, Salary, StartDate, UserName)  
VALUES (4 , '44 Kennedy Way' , NULL , '23453333-21' , 'Kirkland' , '7/23/65' , NULL , 'Helen' , 1 , 'Morgan' , 'Contract Attourney' , 50000 , '1/5/93' , 'helenm')
GO
INSERT INTO tbEmployee(BankID, Address1, Address2, BankAccount, City, DateOfBirth, EndDate, FirstName, Gender, LastName, Position, Salary, StartDate, UserName)  
VALUES (1 , '56 Johnson Blvd' , NULL , '443443442' , 'Helden' , '11/2/66' , NULL , 'Jamie' , 0 , 'Yoder' , 'System Support' , 32000 , '7/2/94' , 'jamiey')
GO
INSERT INTO tbEmployee(BankID, Address1, Address2, BankAccount, City, DateOfBirth, EndDate, FirstName, Gender, LastName, Position, Salary, StartDate, UserName)  
VALUES (4 , '233 Grange Road' , NULL , '477942432' , 'Felixstowe' , '5/1/69' , NULL , 'Martin' , 0 , 'Ashley' , 'Consultant' , 45000 , '4/2/95' , 'martash')
GO
INSERT INTO tbEmployee(BankID, Address1, Address2, BankAccount, City, DateOfBirth, EndDate, FirstName, Gender, LastName, Position, Salary, StartDate, UserName)  
VALUES (3 , '757 Mountain Goat Lane' , NULL , '7362858' , 'Woodley' , '12/28/71' , NULL , 'Phil' , 0 , 'Johnson' , 'Customer Support Engineer' , 32000 , '8/2/94' , 'philjohn')
GO
Print '------------------------------------------------>Completed.'
GO

--Load Sample Dependant Data
PRINT 'INSERTING Dependant Records...'
GO
INSERT INTO tbDependant(DepTypeID, EmployeeID, DateOfBirth, FirstName, Gender, LastName, Status) 
VALUES (1 , 2 , '1/17/59' , 'Lisa' , 1 , 'Fine' , 1)
GO
INSERT INTO tbDependant(DepTypeID, EmployeeID, DateOfBirth, FirstName, Gender, LastName, Status) 
VALUES (2 , 6 , '10/31/89' , 'Mary' , 1 , 'Simpson' , 1)
GO
INSERT INTO tbDependant(DepTypeID, EmployeeID, DateOfBirth, FirstName, Gender, LastName, Status) 
VALUES (1 , 5 , '1/2/64' , 'Sally' , 1 , 'Hargrove-Besser' , 0)
GO
INSERT INTO tbDependant(DepTypeID, EmployeeID, DateOfBirth, FirstName, Gender, LastName, Status) 
VALUES (2 , 5 , '9/4/92' , 'Sam' , 0 , 'Besser' , 1)
GO
INSERT INTO tbDependant(DepTypeID, EmployeeID, DateOfBirth, FirstName, Gender, LastName, Status) 
VALUES (7 , 5 , '9/4/85' , 'Lucy' , 1 , 'Hargrove' , 1)
GO
INSERT INTO tbDependant(DepTypeID, EmployeeID, DateOfBirth, FirstName, Gender, LastName, Status) 
VALUES (1 , 7 , '4/7/65' , 'John' , 0 , 'Morgan' , 0)
GO
INSERT INTO tbDependant(DepTypeID, EmployeeID, DateOfBirth, FirstName, Gender, LastName, Status) 
VALUES (5 , 10 , '11/28/96' , 'Megan' , 1 , 'Johnson' , 0)
GO
Print '------------------------------------------------>Completed.'
GO

--Load Sample Benefit Data
PRINT 'INSERTING Benefit Records...'
GO
INSERT INTO tbBenefit(BnfTaxID, BnfTypeID, Cost, Description, Increment, Maximum, Minimum, Multiples) 
VALUES (3 , 2 , 250000 , 'Life Insurance' , 1000 , 32000 , 10000 , 400)
GO
INSERT INTO tbBenefit(BnfTaxID, BnfTypeID, Cost, Description, Increment, Maximum, Minimum, Multiples) 
VALUES (2 , 1 , 184 , 'Dental Plan' , 3 , 2 , 1 , 4)
GO
INSERT INTO tbBenefit(BnfTaxID, BnfTypeID, Cost, Description, Increment, Maximum, Minimum, Multiples) 
VALUES (3 , 5 , 0 , 'Free Drinks' , 999 , 999 , 0 , 1)
GO
INSERT INTO tbBenefit(BnfTaxID, BnfTypeID, Cost, Description, Increment, Maximum, Minimum, Multiples) 
VALUES (2 , 1 , 247 , 'Healthcare' , 30 , 20 , 10 , 40)
GO
INSERT INTO tbBenefit(BnfTaxID, BnfTypeID, Cost, Description, Increment, Maximum, Minimum, Multiples) 
VALUES (4 , 3 , 500000 , 'Disability Insurance' , 10 , 60 , 40 , 0)
GO
INSERT INTO tbBenefit(BnfTaxID, BnfTypeID, Cost, Description, Increment, Maximum, Minimum, Multiples) 
VALUES (2 , 4 , 1 , 'Pension' , 1 , 9500 , 0 , 15)
GO
INSERT INTO tbBenefit(BnfTaxID, BnfTypeID, Cost, Description, Increment, Maximum, Minimum, Multiples) 
VALUES (2 , 1 , 350 , 'SuperCare' , 30 , 20 , 10 , 40)
GO
INSERT INTO tbBenefit(BnfTaxID, BnfTypeID, Cost, Description, Increment, Maximum, Minimum, Multiples) 
VALUES (2 , 1 , 400 , 'SuperCare Two' , 20 , 20 , 20 , 20)
GO
INSERT INTO tbBenefit(BnfTaxID, BnfTypeID, Cost, Description, Increment, Maximum, Minimum, Multiples) 
VALUES (3 , 5 , 80 , 'Company Car' , 1 , 80 , 0 , 0)
GO
INSERT INTO tbBenefit(BnfTaxID, BnfTypeID, Cost, Description, Increment, Maximum, Minimum, Multiples) 
VALUES (3 , 5 , 20 , 'Gym Membership' , 1 , 20 , 0 , 0)
GO
Print '------------------------------------------------>Completed.'
GO

--Load Sample Employee Benefit Data
PRINT 'INSERTING Employee Benefit Records...'
GO
INSERT INTO tbEmplBnf(BenefitId, CoverageId, EmployeeId, TaxId, Amount) 
VALUES (1 , 4 , 1 , 1 , 30000)
GO
INSERT INTO tbEmplBnf(BenefitId, CoverageId, EmployeeId, TaxId, Amount) 
VALUES (4 , 1 , 2 , 1 , 0)
GO
INSERT INTO tbEmplBnf(BenefitId, CoverageId, EmployeeId, TaxId, Amount) 
VALUES (6 , 1 , 3 , 1 , 11)
GO
INSERT INTO tbEmplBnf(BenefitId, CoverageId, EmployeeId, TaxId, Amount) 
VALUES (2 , 4 , 4 , 2 , 0)
GO
INSERT INTO tbEmplBnf(BenefitId, CoverageId, EmployeeId, TaxId, Amount) 
VALUES (1 , 1 , 4 , 2 , 32000)
GO
INSERT INTO tbEmplBnf(BenefitId, CoverageId, EmployeeId, TaxId, Amount) 
VALUES (5 , 1 , 4 , 2 , 55)
GO
INSERT INTO tbEmplBnf(BenefitId, CoverageId, EmployeeId, TaxId, Amount) 
VALUES (6 , 1 , 4 , 2 , 10)
GO
INSERT INTO tbEmplBnf(BenefitId, CoverageId, EmployeeId, TaxId, Amount) 
VALUES (3 , 1 , 4 , 3 , 0)
GO
INSERT INTO tbEmplBnf(BenefitId, CoverageId, EmployeeId, TaxId, Amount) 
VALUES (2 , 4 , 5 , 2 , 0)
GO
INSERT INTO tbEmplBnf(BenefitId, CoverageId, EmployeeId, TaxId, Amount) 
VALUES (1 , 1 , 5 , 2 , 20000)
GO
INSERT INTO tbEmplBnf(BenefitId, CoverageId, EmployeeId, TaxId, Amount) 
VALUES (5 , 1 , 5 , 2 , 40)
GO
INSERT INTO tbEmplBnf(BenefitId, CoverageId, EmployeeId, TaxId, Amount) 
VALUES (6 , 1 , 5 , 2 , 14)
GO
INSERT INTO tbEmplBnf(BenefitId, CoverageId, EmployeeId, TaxId, Amount) 
VALUES (3 , 1 , 5 , 3 , 0)
GO
Print '------------------------------------------------>Completed.'
GO


PRINT 'Database Load Completed!'
GO
