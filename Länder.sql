CREATE TABLE [Länder] (
  [LandID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Ländercode] VARCHAR (2),
  [Landname] VARCHAR (50),
  [PLZ_Format] VARCHAR (50),
  [Bemerkung] VARCHAR (100)
)
