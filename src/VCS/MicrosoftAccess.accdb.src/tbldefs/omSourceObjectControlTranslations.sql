CREATE TABLE [omSourceObjectControlTranslations] (
  [Id] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [SourceObjectControlId] LONG,
  [LanguageId] LONG,
  [Default] VARCHAR (255),
  [Short] VARCHAR (255),
  [Long] LONGTEXT,
  [CreateDate] DATETIME,
  [LastUsedDate] DATETIME
)
