
/****** Object:  Table [dbo].[Beverages]    Script Date: 10/29/01 14:56:26 ******/
if exists (select * from sysobjects where id = object_id(N'[dbo].[Beverages]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Beverages]
GO

/****** Object:  Table [dbo].[Beverages]    Script Date: 10/29/01 14:56:27 ******/
CREATE TABLE [dbo].[Beverages] (
	           [iDrinkID] [int] IDENTITY (1, 1) NOT NULL ,
	           [vcDrink] [varchar] (50) NOT NULL ,
	           [mPrice] [money] NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Beverages] WITH NOCHECK ADD 
	CONSTRAINT [PK_tBeverages] PRIMARY KEY  NONCLUSTERED 
	(
		[iDrinkID]
	)  ON [PRIMARY] 
GO
/*
SET NOCOUNT ON
INSERT Beverages (vcDrink, mPrice) VALUES ('CocaCola', .50)
INSERT Beverages (vcDrink, mPrice) VALUES ('DietCoke', .50)
INSERT Beverages (vcDrink, mPrice) VALUES ('Pepsi', .50)
INSERT Beverages (vcDrink, mPrice) VALUES ('DietPepsi', .50)
INSERT Beverages (vcDrink, mPrice) VALUES ('MountainDew', .50)
SET NOCOUNT OFF
SELECT * FROM Beverages
DELETE FROM Beverages WHERE iDrinkID > 5   
SELECT * FROM Beverages
-- Truncate table Beverages
*/



/****** Object:  Table [dbo].[TranLog]    Script Date: 10/29/01 14:57:36 ******/
if exists (select * from sysobjects where id = object_id(N'[dbo].[TranLog]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TranLog]
GO

/****** Object:  Table [dbo].[TranLog]    Script Date: 10/29/01 14:57:36 ******/
CREATE TABLE [dbo].[TranLog] (
	[iTranID] [int] IDENTITY (1, 1) NOT NULL ,
	[iItemID] [int] NULL ,
	[iQuantity] [int] NULL ,
	[mAmount] [money] NULL ,
	[vcTranComment] [varchar] (50) NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[TranLog] WITH NOCHECK ADD 
	CONSTRAINT [PK_TranLog] PRIMARY KEY  NONCLUSTERED 
	(
		[iTranID]
	)  ON [PRIMARY] 
GO
/*
SET NOCOUNT ON
SELECT * FROM TranLog

--INSERT TranLog (iItemID, iQuantity, mAmount, vcTranComment) VALUES (1, -1 , .50, 'Purchase')          --Purchase

INSERT TranLog (iItemID, iQuantity, mAmount, vcTranComment) VALUES (1, 10 , 0, 'Stock 10 CocaCola')   --Stock
INSERT TranLog (iItemID, iQuantity, mAmount, vcTranComment) VALUES (2, 10 , 0, 'Stock 10 DietCoke')   --Stock
INSERT TranLog (iItemID, iQuantity, mAmount, vcTranComment) VALUES (3, 10 , 0, 'Stock 10 Pepsi')      --Stock
INSERT TranLog (iItemID, iQuantity, mAmount, vcTranComment) VALUES (4, 10 , 0, 'Stock 10 DietPepsi')  --Stock
INSERT TranLog (iItemID, iQuantity, mAmount, vcTranComment) VALUES (5, 10 , 0, 'Stock 10 MountainDew')--Stock
SELECT * FROM TranLog
SET NOCOUNT OFF
-- Truncate table TranLog
*/


/****** Display Tables ****/
SET NOCOUNT ON
SELECT * FROM Beverages   -- Truncate table Beverages
SELECT * FROM TranLog     -- Truncate table TranLog
SET NOCOUNT OFF

SELECT SUM(iQuantity), 'CocaCola in stock' FROM TranLog WHERE iItemID = 1
SELECT SUM(iQuantity), 'DietCoke in stock' FROM TranLog WHERE iItemID = 2
SELECT SUM(iQuantity), 'Pepsi in stock' FROM TranLog WHERE iItemID = 3
SELECT SUM(iQuantity), 'DietPepsi in stock' FROM TranLog WHERE iItemID = 4
SELECT SUM(iQuantity), 'MountainDew in stock' FROM TranLog WHERE iItemID = 5






