SELECT
  omSourceObjectControlTranslations.Id,
  omSourceObjectControls.SourceObjectId,
  omSourceObjects.Name AS SourceObjectName,
  omSourceObjectControls.ControlName,
  omSourceObjectControlTranslations.LanguageId,
  omLanguages.Name AS LanguageName,
  omSourceObjectControls.ControlDefault,
  omSourceObjectControlTranslations.Default,
  omSourceObjectControlTranslations.Short,
  omSourceObjectControlTranslations.Long
FROM
  (
    (
      omSourceObjectControlTranslations
      INNER JOIN omSourceObjectControls ON omSourceObjectControlTranslations.SourceObjectControlId = omSourceObjectControls.Id
    )
    INNER JOIN omSourceObjects ON omSourceObjectControls.SourceObjectId = omSourceObjects.Id
  )
  INNER JOIN omLanguages ON omSourceObjectControlTranslations.LanguageId = omLanguages.Id
WHERE
  (
    (
      (omLanguages.Name) Like "*" & Forms!omSourceObjectControlTranslation_List!txtSearch & "*"
    )
  )
  Or (
    (
      (omSourceObjects.Name) Like "*" & Forms!omSourceObjectControlTranslation_List!txtSearch & "*"
    )
  )
  Or (
    (
      (
        omSourceObjectControls.ControlName
      ) Like "*" & Forms!omSourceObjectControlTranslation_List!txtSearch & "*"
    )
  )
  Or (
    (
      (
        omSourceObjectControls.ControlDefault
      ) Like "*" & Forms!omSourceObjectControlTranslation_List!txtSearch & "*"
    )
  )
  Or (
    (
      (
        omSourceObjectControlTranslations.Default
      ) Like "*" & Forms!omSourceObjectControlTranslation_List!txtSearch & "*"
    )
  )
  Or (
    (
      (
        omSourceObjectControlTranslations.Short
      ) Like "*" & Forms!omSourceObjectControlTranslation_List!txtSearch & "*"
    )
  )
  Or (
    (
      (
        omSourceObjectControlTranslations.Long
      ) Like "*" & Forms!omSourceObjectControlTranslation_List!txtSearch & "*"
    )
  )
ORDER BY
  omSourceObjects.Name,
  omSourceObjectControls.ControlName,
  omLanguages.Name;
