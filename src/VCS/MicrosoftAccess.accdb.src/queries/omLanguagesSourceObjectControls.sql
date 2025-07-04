SELECT
  omLanguages.Id AS LanguageId,
  omSourceObjectControls.Id AS SourceObjectControlId,
  omSourceObjectControls.ControlDefault
FROM
  omLanguages,
  omSourceObjectControls;
