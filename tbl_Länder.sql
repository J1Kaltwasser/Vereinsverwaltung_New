CREATE TABLE [tbl_Länder] (
  [Land] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Ländercode] VARCHAR (4),
  [IsEU] BIT,
  [IsAktiv] BIT
)
