CREATE TABLE [omSourceObjectControls] (
  [Id] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [SourceObjectId] LONG,
  [ControlId] LONG,
  [ControlName] VARCHAR (255),
  [ControlTypeId] LONG,
  [ControlDefault] VARCHAR (255),
  [CreateDate] DATETIME,
  [LastUsedDate] DATETIME
)
