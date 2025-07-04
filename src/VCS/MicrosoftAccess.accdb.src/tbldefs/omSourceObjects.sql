CREATE TABLE [omSourceObjects] (
  [Id] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ObjectTypeId] LONG,
  [Name] VARCHAR (255),
  [CreateDate] DATETIME,
  [LastUsedDate] DATETIME
)
