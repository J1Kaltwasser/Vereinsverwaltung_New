CREATE TABLE [tbl_AuditLog] (
  [AuditID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Tabelle] VARCHAR (50),
  [RecordID] LONG,
  [Feld] VARCHAR (100),
  [AlterWert] LONGTEXT,
  [NeuerWert] LONGTEXT,
  [ÄnderungsDatum] DATETIME,
  [BenutzerName] VARCHAR (50),
  [Aktion] VARCHAR (20)
)
