CREATE TABLE [Hierarchische Stellungen] (
  [StellungsID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Stellungsbezeichnung] VARCHAR (50),
  [Ebene] LONG
)
