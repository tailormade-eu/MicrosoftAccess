INSERT INTO omSourceObjectControlTranslations (
  LanguageId, SourceObjectControlId,
  [Default], [Short], [Long], CreateDate,
  LastUsedDate
)
SELECT
  omLanguagesSourceObjectControls.LanguageId,
  omLanguagesSourceObjectControls.SourceObjectControlId,
  omLanguagesSourceObjectControls.ControlDefault,
  omLanguagesSourceObjectControls.ControlDefault,
  omLanguagesSourceObjectControls.ControlDefault,
  Now() AS CreateDate,
  Now() AS LastUsedDate
FROM
  omLanguagesSourceObjectControls
  LEFT JOIN omSourceObjectControlTranslations ON (
    omLanguagesSourceObjectControls.SourceObjectControlId = omSourceObjectControlTranslations.SourceObjectControlId
  )
  AND (
    omLanguagesSourceObjectControls.LanguageId = omSourceObjectControlTranslations.LanguageId
  )
WHERE
  (
    (
      (
        omSourceObjectControlTranslations.Id
      ) Is Null
    )
  );
