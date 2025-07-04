CREATE TABLE [omControls] (
  [Id] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ControlTypeId] LONG,
  [Name] VARCHAR (255),
  [CreateDate] DATETIME,
  [LastUsedDate] DATETIME
)
