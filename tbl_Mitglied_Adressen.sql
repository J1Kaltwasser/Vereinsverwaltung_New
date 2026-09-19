CREATE TABLE [tbl_Mitglied_Adressen] (
  [AdressID] AUTOINCREMENT CONSTRAINT [AdressID] UNIQUE CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [MitgliedID] LONG CONSTRAINT [MitgliedID] UNIQUE,
  [Adresstyp] VARCHAR (50),
  [Straße] VARCHAR (100),
  [Hausnummer] VARCHAR (20),
  [PLZ] VARCHAR (10),
  [Stadt] VARCHAR (50),
  [Bundesland] VARCHAR (50),
  [Land] VARCHAR (50),
  [GültigAb] DATETIME,
  [GültigBis] DATETIME,
  [IsPrimär] BIT,
  [Anmerkungen] LONGTEXT,
  [Created_Date] DATETIME,
  [Modified_Date] DATETIME
)
