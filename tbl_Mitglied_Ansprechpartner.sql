CREATE TABLE [tbl_Mitglied_Ansprechpartner] (
  [AnsprechpartnerID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [MitgliedID] LONG,
  [Anrede] VARCHAR (20),
  [Vorname] VARCHAR (50),
  [Nachname] VARCHAR (50),
  [Titel] VARCHAR (30),
  [Email] VARCHAR (100),
  [Telefon] VARCHAR (30),
  [Mobil] VARCHAR (30),
  [Fax] VARCHAR (30),
  [IsPrimär] BIT,
  [Created_Date] DATETIME,
  [Modified_Date] DATETIME
)
