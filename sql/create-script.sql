--Author: Ryan Beckett
--© 2018
CREATE DATABASE churches
--ON PRIMARY
--(NAME = churches_dat,
--    FILENAME = 'c:\db\data\churches_dat.mdf',
--    SIZE = 10MB,
--    MAXSIZE = 4000MB,
--    FILEGROWTH = 10MB ),
--FILEGROUP churches_ftdata
--(NAME = churches_ftdata,
--	FILENAME = 'c:\db\churches_ftdata.ndf',
--	SIZE = 10MB,
--	MAXSIZE = 400MB,
--	FILEGROWTH = 1MB
--)
--LOG ON
--(NAME = churches_log,
--    FILENAME = 'c:\db\log\churches_log.ldf',
--    SIZE = 5MB,
--    MAXSIZE = 2000MB,
--    FILEGROWTH = 5MB )
GO
--ensure SQL Server is always available
exec sp_dboption N'churches', N'autoclose', N'false'
GO
--no logfiles needed
ALTER DATABASE churches
SET RECOVERY SIMPLE
GO
USE churches
IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled')) 
    begin 
	    EXEC sp_fulltext_database @action = 'enable'
    end
GO
-- SQL Server full text catalogs
-- makes looking up users on the admin side easy
-- not to be used for user-facing applications
CREATE FULLTEXT CATALOG ChurchName --IN PATH 'c:\db\' --ON FILEGROUP churches_ftdata
GO
CREATE FULLTEXT CATALOG Email --IN PATH 'c:\db\' --ON FILEGROUP churches_ftdata
GO
CREATE FULLTEXT CATALOG UserName --IN PATH 'c:\db\' --ON FILEGROUP churches_ftdata
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Churches](
	[ChurchID] [int] IDENTITY(1,1) NOT NULL,
	[ChurchName] [nvarchar](73) NULL,
	[Denomination] [nvarchar](101) NULL,
	[Address] [nvarchar](73) NULL,
	[Mail] [nvarchar](129) NULL,
	[City] [nvarchar](51) NULL,
	[Province] [nvarchar](51) NULL,
	[Region] [nvarchar](51) NULL,
	[PostalCode] [nvarchar](11) NULL,
	[Phone] [nvarchar](51) NULL,
	[Fax] [nvarchar](51) NULL,
	[ContactPerson] [nvarchar](51) NULL,
	[Email] [nvarchar](151) NULL,
	[Website] [nvarchar](151) NULL,
	[DenomWebsite] [nvarchar](151) NULL,
	[LastModified] [datetime] NULL,
CONSTRAINT [PK_ChurchID] PRIMARY KEY CLUSTERED 
(
	[ChurchID] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
UPDATE Churches SET LastModified = '2018-05-11 00:00:01';
GO
CREATE FULLTEXT INDEX ON [Churches]
([ChurchName] LANGUAGE [English])
KEY INDEX [PK_ChurchID] ON [ChurchName]
WITH CHANGE_TRACKING AUTO
GO
ALTER FULLTEXT INDEX ON [Churches] START FULL POPULATION
GO
CREATE TABLE [User](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[GUID] [uniqueidentifier] ROWGUIDCOL  NOT NULL CONSTRAINT [DF_User_GUID]  DEFAULT (newid()),
	[Name] [nvarchar](50) NOT NULL,
	[Email] [nvarchar](50) NOT NULL,
	[Password] [varbinary](50) NOT NULL,
	[Created] [datetime] NOT NULL CONSTRAINT [DF_User_Created]  DEFAULT (getdate()),
	[Modified] [datetime] NOT NULL CONSTRAINT [DF_User_Modified]  DEFAULT (getdate()),
	[Confirmed] [bit] NOT NULL CONSTRAINT [DF_User_Confirmed]  DEFAULT ((0)),
	[Salt] [binary](4) NOT NULL,
 CONSTRAINT [PK_UserID] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
CREATE FULLTEXT INDEX ON [User](
[Email] LANGUAGE [English], 
[Name] LANGUAGE [English])
KEY INDEX [PK_UserID] ON [UserName]
WITH CHANGE_TRACKING AUTO
GO
ALTER FULLTEXT INDEX ON [User] START FULL POPULATION
GO
CREATE TABLE [ChurchEmail](
	[ChurchID] [int] NOT NULL,
	[Email] [nvarchar](150) NOT NULL,
 CONSTRAINT [PK_EmailID] PRIMARY KEY CLUSTERED 
(
	[ChurchID] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
CREATE FULLTEXT INDEX ON [ChurchEmail](
[Email] LANGUAGE [English])
KEY INDEX [PK_EmailID] ON [Email]
WITH CHANGE_TRACKING AUTO
GO
CREATE TABLE [IPTable](
	[ConnID] numeric(20,0) NOT NULL,
	[Count] int NOT NULL,
 CONSTRAINT [PK_IPTableID] PRIMARY KEY CLUSTERED 
(
	[ConnID] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER FULLTEXT INDEX ON [ChurchEmail] START FULL POPULATION
GO
CREATE PROCEDURE [dbo].[spConfirmUser]
(	@guid uniqueidentifier	)
AS
	UPDATE [dbo].[User]
	SET
	[dbo].[User].[Confirmed] = 1 WHERE 
	[dbo].[User].[GUID] = @guid AND [dbo].[User].[Confirmed] <> 1
GO
CREATE PROCEDURE [dbo].[spChurchEmail]
(	@email nvarchar(150),
	@churchid int	)
AS
	SELECT [dbo].[ChurchEmail].[Email], [dbo].[ChurchEmail].[ChurchID] FROM [dbo].[ChurchEmail] WHERE 
	CONTAINS (Email, @email) AND (ChurchID = @churchid)
GO
CREATE PROCEDURE [spProvRegCityDenomChurchIDs]
(	@province nvarchar(50),
	@region nvarchar(50),
	@city nvarchar(50),
	@denomination nvarchar(100) )
AS SELECT ChurchName, Address, Mail, PostalCode, ChurchID FROM Churches WHERE
(Province = @province) AND (Region = @region) AND (City = @city) AND
(Denomination = @denomination)
GO
CREATE PROCEDURE [dbo].[spAddChurch]
(	@churchname nvarchar(72),
	@denomination nvarchar(100),
	@address nvarchar(72),
	@mail nvarchar(128),
	@city nvarchar(50),
	@province nvarchar(50),
	@region nvarchar(50),
	@postalcode nvarchar(10),
	@phone nvarchar(50),
	@fax nvarchar(50),
	@contactperson nvarchar(50),
	@email nvarchar(150),
	@website nvarchar(150),
	@denomwebsite nvarchar(150),
	@churchid int OUTPUT	)
AS
	BEGIN
		INSERT INTO [dbo].[Churches]
		(
			ChurchName, Denomination, Address, Mail, City, Province, Region, PostalCode, Phone, Fax, ContactPerson, Email, Website, DenomWebsite
		)
		VALUES
		(
			@churchname, @denomination, @address, @mail, @city, @province, @region, @postalcode, @phone, @fax, @contactperson, @email, @website, @denomwebsite
		)
		SELECT @churchid = [dbo].[Churches].[ChurchID] FROM [dbo].[Churches] WHERE 
		[dbo].[Churches].[ChurchID] = @@identity
	END
GO
CREATE PROCEDURE [dbo].[spUpdChurch]
(	@churchid int,
	@churchname nvarchar(72),
	@denomination nvarchar(100),
	@address nvarchar(72),
	@mail nvarchar(128),
	@city nvarchar(50),
	@province nvarchar(50),
	@region nvarchar(50),
	@postalcode nvarchar(10),
	@phone nvarchar(50),
	@fax nvarchar(50),
	@contactperson nvarchar(50),
	@email nvarchar(150),
	@website nvarchar(150),
	@denomwebsite nvarchar(150)	)
AS
	BEGIN
		UPDATE [dbo].[Churches] SET [dbo].[Churches].[ChurchName] = @churchname, [dbo].[Churches].[Denomination] = @denomination, 
[dbo].[Churches].[Address] = @address, [dbo].[Churches].[Mail] = @mail, [dbo].[Churches].[City] = @city, [dbo].[Churches].[Province] = @province, 
[dbo].[Churches].[Region] = @region, [dbo].[Churches].[PostalCode] = @postalcode, [dbo].[Churches].[Phone] = @phone, [dbo].[Churches].[Fax] = @fax, 
[dbo].[Churches].[ContactPerson] = @contactperson, [dbo].[Churches].[Email] = @email, [dbo].[Churches].[Website] = @website, [dbo].[Churches].[DenomWebsite] = @denomwebsite 
		WHERE [dbo].[Churches].[ChurchID] = @churchid
	END
GO
CREATE PROCEDURE [dbo].[spAddChurchEmail]
(	@churchid int,
	@email nvarchar(150)	)
AS
	BEGIN
		INSERT INTO [dbo].[ChurchEmail]
		(
			ChurchID, Email
		)
		VALUES
		(
			@churchid, @email
		)
	END
GO
CREATE PROCEDURE [dbo].[spUpdChurchEmail]
(	@churchid int,
	@email nvarchar(150)	)
AS
	BEGIN
		UPDATE [dbo].[ChurchEmail] SET [dbo].[ChurchEmail].[Email] = @email WHERE [dbo].[ChurchEmail].[ChurchID] = @churchid
	END
GO
CREATE PROCEDURE [dbo].[spRmvChurch]
(	@churchid int	)
AS
	BEGIN
		DELETE FROM [dbo].[Churches] WHERE [dbo].[Churches].[ChurchID] = @churchID
	END
GO
CREATE PROCEDURE [spAllProvs]
AS SELECT DISTINCT Province FROM Churches
ORDER BY Province
GO
CREATE PROCEDURE [spProvRegions]
(	@province nvarchar(50)		)
AS SELECT DISTINCT Region FROM Churches WHERE (Province = @province)
GO
CREATE PROCEDURE [spProvRegCities]
(	@province nvarchar(50),
	@region nvarchar(50)		)
AS SELECT DISTINCT City FROM Churches
WHERE (Province = @province) AND (Region = @region)
GO
CREATE PROCEDURE [spProvRegCityDenoms]
(	@province nvarchar(50),
	@region nvarchar(50),
	@city nvarchar(50)			)
AS SELECT DISTINCT Denomination FROM Churches 
WHERE (Province = @province) AND (Region = @region) 
AND (City = @city)
GO
CREATE PROCEDURE [spProvRegCityDenomChurches]
(	@province nvarchar(50),
	@region nvarchar(50),
	@city nvarchar(50),
	@denomination nvarchar(100)	)
AS SELECT ChurchName, Address, Mail, PostalCode, LastModified FROM Churches WHERE 
(Province = @province) AND (Region = @region) AND (City = @city) AND (Denomination = @denomination) 
ORDER BY ChurchName
GO
CREATE PROCEDURE [spChurchDetails](@churchid int)
AS SELECT     ChurchName, Denomination, Address, Mail, City, Province, Region, PostalCode, Phone, Fax, ContactPerson, Email, Website, DenomWebsite
FROM         Churches
WHERE     (ChurchID = @churchid)
GO
CREATE PROCEDURE [spProvRegCityDenomChurch]
(	@province nvarchar(50),
	@region nvarchar(50),
	@city nvarchar(50),
	@denomination nvarchar(100),
	@churchname nvarchar(72)	)
AS SELECT     ChurchName, Denomination, Address, Mail, City, Province, Region, PostalCode, Phone, Fax, ContactPerson, Email, Website, DenomWebsite
FROM         Churches
WHERE     (Province = @province) AND (Region = @region) AND (City = @city) AND
(Denomination = @denomination) AND CONTAINS(ChurchName, @churchname)
GO
CREATE PROCEDURE [spRegCityDenomChurches]
(	@region nvarchar(50),
	@city nvarchar(50),
	@denomination nvarchar(100)	)
AS SELECT ChurchName, Address, Mail, PostalCode, LastModified FROM Churches WHERE 
(Region = @region) AND (City = @city) AND (Denomination = @denomination)
GO
CREATE PROCEDURE [spRegCityChurches]
(	@region nvarchar(50),
	@city nvarchar(50)	)
AS SELECT ChurchName, Address, Mail, PostalCode, LastModified FROM Churches WHERE 
(Region = @region) AND (City = @city) 
GO
CREATE PROCEDURE [spRegDenomChurches]
(	@region nvarchar(50),
	@denomination nvarchar(100)	)
AS SELECT ChurchName, Address, Mail, PostalCode, LastModified FROM Churches WHERE 
(Region = @region) AND (Denomination = @denomination)
GO
CREATE PROCEDURE [spRegChurches]
(	@region nvarchar(50)	)
AS SELECT ChurchName, Address, Mail, PostalCode, City, LastModified FROM Churches 
WHERE (Region = @region) ORDER BY City
GO
CREATE PROCEDURE [spProvRegDenomCities]
(	@province nvarchar(50),
	@region nvarchar(50), 
	@denomination nvarchar(100)		)
AS SELECT DISTINCT City FROM Churches 
WHERE (Province = @province) AND (Region = @region) AND (Denomination = @denomination) 
GO
CREATE PROCEDURE [spProvRegDenomChurchNameChurches]
(	@province nvarchar(50),
	@region nvarchar(50),
	@denomination nvarchar(100),
	@churchname nvarchar(72)	)
AS SELECT ChurchName, Address, Mail, PostalCode, City, ChurchID, LastModified FROM Churches WHERE
(Province = @province) AND (Region = @region) AND(Denomination = @denomination) 
AND CONTAINS(ChurchName, @churchname)
GO
CREATE PROCEDURE [spProvRegCityDenomChurchNameChurches]
(	@province nvarchar(50),
	@region nvarchar(50),
	@city nvarchar(50),
	@denomination nvarchar(100),
	@churchname nvarchar(72)	)
AS SELECT ChurchName, Address, Mail, PostalCode, ChurchID, LastModified FROM Churches WHERE
(Province = @province) AND (Region = @region) AND (City = @city) AND
(Denomination = @denomination) AND CONTAINS(ChurchName, @churchname)
GO
CREATE PROCEDURE [spProvRegCityChurches]
(	@province nvarchar(50),
	@region nvarchar(50),
	@city nvarchar(50)	)
AS SELECT ChurchName, Address, Mail, PostalCode, Denomination, LastModified FROM Churches 
WHERE (Province = @province) AND (Region = @region) AND (City = @city) 
ORDER BY ChurchName
GO
CREATE PROCEDURE [spProvRegDenomChurches]
(	@province nvarchar(50),
	@region nvarchar(50),
	@denomination nvarchar(100)	)
AS SELECT ChurchName, Address, PostalCode, City, LastModified FROM Churches 
WHERE (Province = @province) AND (Region = @region) AND (Denomination = @denomination) 
ORDER BY ChurchName
GO
CREATE PROCEDURE [spProvRegCityChurchNameChurches]
(	@province nvarchar(50),
	@region nvarchar(50),
	@city nvarchar(50),
	@churchname nvarchar(72)	)
AS SELECT ChurchName, Address, Mail, PostalCode, ChurchID, LastModified FROM Churches WHERE 
(Province = @province) AND (Region = @region) AND (City = @city) AND CONTAINS(ChurchName, @churchname)
GO
CREATE PROCEDURE [spProvRegChurches]
(	@province nvarchar(50),
	@region nvarchar(50)	)
AS SELECT ChurchName, Address, Mail, PostalCode, City, LastModified FROM Churches 
WHERE (Province = @province) AND (Region = @region) ORDER BY City 
GO
CREATE PROCEDURE [spProvRegChurchNameChurches]
(	@province nvarchar(50),
	@region nvarchar(50),
	@churchname nvarchar(72)	)
AS SELECT ChurchName, Address, PostalCode, City, ChurchID, LastModified FROM Churches WHERE 
(Province = @province) AND (Region = @region) AND CONTAINS(ChurchName, @churchname)
ORDER BY City
GO
CREATE PROCEDURE [spProvDenomCities]
(	@province nvarchar(50),
	@denomination nvarchar(100)		)
AS SELECT DISTINCT City FROM Churches 
WHERE (Province = @province) AND (Denomination = @denomination) 
GO
CREATE PROCEDURE [spProvDenomChurches]
(	@province nvarchar(50),
	@denomination nvarchar(100)	)
AS SELECT ChurchName, Address, Mail, PostalCode, City, LastModified FROM Churches 
WHERE (Province = @province) AND (Denomination = @denomination)
GO
CREATE PROCEDURE [spProvDenomChurchNameChurches]
(	@province nvarchar(50),
	@denomination nvarchar(100),
	@churchname nvarchar(72)	)
AS SELECT ChurchName, Address, Mail, PostalCode, City, ChurchID, LastModified FROM Churches 
WHERE (Province = @province) AND 
(Denomination = @denomination) AND CONTAINS(ChurchName, @churchname)
GO
CREATE PROCEDURE [spProvCityDenomChurches]
(	@province nvarchar(50),
	@city nvarchar(50),
	@denomination nvarchar(100)	)
AS SELECT ChurchName, Address, Mail, PostalCode, LastModified FROM Churches 
WHERE (Province = @province) AND (City = @city) AND (Denomination = @denomination)
GO
CREATE PROCEDURE [spProvCityDenomChurchNameChurches]
(	@province nvarchar(50),
	@city nvarchar(50),
	@denomination nvarchar(100),
	@churchname nvarchar(73)	)
AS SELECT ChurchName, Address, Mail, PostalCode, ChurchID, LastModified FROM Churches WHERE
(Province = @province) AND (City = @city) AND(Denomination = @denomination) 
AND CONTAINS(ChurchName, @churchname)
GO
CREATE PROCEDURE [spProvCityChurches]
(	@province nvarchar(50),
	@city nvarchar(50)			)
AS SELECT ChurchName, Address, Mail, PostalCode, Denomination, LastModified FROM Churches 
WHERE (Province = @province) AND (City = @city)
GO
CREATE PROCEDURE [spProvCityChurchNameChurches]
(	@province nvarchar(50),
	@city nvarchar(50),
	@churchname nvarchar(72)	)
AS SELECT  ChurchName, Address, Mail, PostalCode, ChurchID, LastModified FROM Churches 
WHERE (Province = @province) AND (City = @city) AND CONTAINS(ChurchName, @churchname)
GO
CREATE PROCEDURE [spProvChurches]
(	@province nvarchar(50)	)
AS SELECT ChurchName, Address, Mail, PostalCode, City, Region, LastModified 
FROM Churches WHERE (Province = @province) ORDER BY Region
GO
CREATE PROCEDURE [spProvChurchNameChurches]
(	@province nvarchar(50),
	@churchname nvarchar(72)	)
AS SELECT ChurchName, Address, Mail, PostalCode, City, ChurchID, LastModified FROM Churches 
WHERE (Province = @province) AND CONTAINS(ChurchName, @churchname)
GO
CREATE PROCEDURE [spDenomChurches]
(	@denomination nvarchar(100)	)
AS SELECT ChurchName, Address, Mail, PostalCode, City, Province, LastModified FROM 
Churches WHERE (Denomination = @denomination) ORDER BY Province
GO
CREATE PROCEDURE [spCityChurches]
( @city nvarchar(50) )
AS SELECT ChurchName, Address, Mail, PostalCode, LastModified FROM Churches WHERE 
(City = @city) ORDER BY ChurchName
GO
CREATE PROCEDURE [spChurchNameProvChurches]
(@province nvarchar(50),
@churchname nvarchar(72))
AS SELECT ChurchName, Address, PostalCode, City, Denomination FROM Churches
WHERE (Province = @province) AND CONTAINS(ChurchName, @churchname)
ORDER BY City
GO
CREATE PROCEDURE [spChurchNameChurches](@churchname nvarchar(72))
AS SELECT ChurchName, Address, PostalCode, City, Province, Denomination, LastModified FROM Churches
WHERE CONTAINS(ChurchName, @churchname)
ORDER BY Province
GO
CREATE PROCEDURE [spCityDenomChurches](
	 @city nvarchar(50),
	 @denomination nvarchar(100)	)
AS SELECT ChurchName, Address, Mail, PostalCode, LastModified FROM 
Churches WHERE (City = @city) AND (Denomination = @denomination) ORDER BY Province
GO
CREATE PROCEDURE spAddIP
(	@connectionid numeric(20,0),
	@newcount int	)
AS INSERT INTO IPTable (ConnID, Count) VALUES (@connectionid, @newcount)
GO
CREATE PROCEDURE spUpdateIP
(	@connectionid numeric(20,0),
	@newcount int	)
AS UPDATE IPTable SET [Count] = @newcount WHERE ConnID = @connectionid
GO
CREATE PROCEDURE spGetIPCount
(   @connectionid numeric(20,0)   )
AS SELECT [Count] FROM IPTable WHERE ConnID = @connectionid;
GO
