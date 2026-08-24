# Translation coverage

Written by `scripts/devtools/translation-coverage.R` on 2026-08-24 10:24.

Workbook read: `trads/designer_translations.xlsx`.

**Missing** means the product asks for the tag and no table row carries it. The screen then shows the tag itself. **Dead** means a row nobody asks for.

## Summary

| table | rows | missing | dead | oracle |
| --- | --- | --- | --- | --- |
| t_tradllshapes | 10 | see the message list | 2 | VBA string literals |
| t_tradllmsg | 175 | see the message list | 58 | VBA string literals |
| t_tradllforms | 125 | 0 | 29 | control names in `.mock/forms/designer` |
| t_tradllribbon | 38 | 0 | 11 | getLabel ids in the linelist ribbon templates |
| t_tradmsg | 78 | 0 | 49 | VBA literals + getLabel ids in the designer ribbon |
| t_tradrange | 14 | 2 to check | 1 | defined names in designer.xlsb |
| t_tradshape | 13 | 0 | 8 | shape names in designer.xlsb |
| t_traddrop | 1 | 0 | 1 | defined names in designer.xlsb |

Message tags the code asks for and no table carries: **0**. They are listed on their own below, because a string literal does not say which of the two sheets should carry it.

## Missing: message tags

_Nothing._

## Missing: form controls

_Nothing._

## Missing: linelist ribbon

_Nothing._

## Missing: designer ribbon

_Nothing._

## Candidates: designer worksheet labels

A defined name spelled like a label that the table does not carry. The file does not say which sheet a name points at, and only Main is translated, so open the workbook before adding a row: a name on an internal sheet such as `__pass` is not a label anybody reads.

- `RNG_LabPrivateKey`
- `RNG_LabPublicKey`

## Missing: designer worksheet buttons

_Nothing._

## Dead rows: linelist messages (t_tradllmsg)

- `INSTSHEET`
- `MSG_AnaSheet`
- `MSG_Analysis`
- `MSG_BadLLName`
- `MSG_ChooseDir`
- `MSG_ChooseFile`
- `MSG_CloseImportFile`
- `MSG_CustomTableSheet`
- `MSG_Debug`
- `MSG_DesHidden`
- `MSG_ErrClearHistoric`
- `MSG_ErrDebug`
- `MSG_ErrExportData`
- `MSG_ErrExportGeo`
- `MSG_ErrExportHistoricGeo`
- `MSG_ErrExportPath`
- `MSG_ErrHistoricGeo`
- `MSG_ErrShowHide`
- `MSG_ExcelFile`
- `MSG_FinishImportHistoricGeo`
- `MSG_FinishedClear`
- `MSG_From`
- `MSG_Hidden`
- `MSG_Hide`
- `MSG_HideAllOptional`
- `MSG_ImportDone`
- `MSG_ImportGeoDone`
- `MSG_InvalidFormula`
- `MSG_M`
- `MSG_MaxData`
- `MSG_MinData`
- `MSG_NewPass`
- `MSG_NoExport`
- `MSG_NoImportDone`
- `MSG_NoImportGeoDone`
- `MSG_NotModifyHeader`
- `MSG_PathImport`
- `MSG_PathTooLong`
- `MSG_PrintSheet`
- `MSG_Protect`
- `MSG_ProvidePassword`
- `MSG_RangeAna`
- `MSG_SectionNotExit`
- `MSG_SelectSection`
- `MSG_Show`
- `MSG_ShowAllOptional`
- `MSG_ShowHoriz`
- `MSG_ShowMandatory`
- `MSG_ShowVerti`
- `MSG_Shown`
- `MSG_To`
- `MSG_UnableShowHide`
- `MSG_UnableToAgg`
- `MSG_WaitPlease`
- `MSG_WrongPassword`
- `MSG_exportErrHandData`
- `MSG_exportErrHandWrite`
- `MSG_per`

## Dead rows: linelist shapes (t_tradllshapes)

- `SHP_Debug`
- `SHP_Reset`

## Dead rows: linelist forms (t_tradllforms)

- `CHK_ExportMigData`
- `CMD_ApplyFilter`
- `CMD_Export1`
- `CMD_Export10`
- `CMD_Export2`
- `CMD_Export3`
- `CMD_Export4`
- `CMD_Export5`
- `CMD_Export6`
- `CMD_Export7`
- `CMD_Export8`
- `CMD_Export9`
- `CMD_GeoClearHisto`
- `CMD_ImportData`
- `CMD_ImportGeo`
- `CMD_ImportGeoHistoric`
- `CMD_RemoveFilter`
- `CMD_RenameFilter`
- `CMD_Retour`
- `CMD_SaveToFilter`
- `F_Filters`
- `F_ShowHideLL`
- `F_ShowHidePrint`
- `FRM_Geo`
- `LBL_EmptySection`
- `LBL_FilterName`
- `LBL_ImpRepSheets`
- `LBL_PrintVarName`
- `LBL_SheetName`

## Dead rows: linelist ribbon (t_tradllribbon)

- `adminTab`
- `analysisTab`
- `btnAddCustomFilt`
- `btnApplyFilt`
- `btnCustomFilt`
- `btnDebug`
- `customGroupAdmin`
- `customGroupTabFilt`
- `customGroupTabGeo`
- `customPivotTab`
- `dataEntryTab`

## Dead rows: designer messages and ribbon (t_tradmsg)

The designer workbook's own document modules live inside `designer.xlsb` and nowhere in `src`, so a tag one of them asks for on its own reads as dead here. Check the file before deleting a row.

- `customGroupAdvanced`
- `MSG_Exists`
- `NoteText_Forbidden_Caracteres`
- `MSG_HListVList`
- `MSG_BuildAna`
- `MSG_ConfCancel`
- `MSG_PathDic`
- `MSG_VeriFichGeo`
- `MSG_PathLL`
- `MSG_NetoPrec`
- `MSG_CloseDic`
- `MSG_BuildLL`
- `MSG_Title_Dictionnary`
- `MSG_AlreadyOpen`
- `MSG_Question`
- `MSG_Fini`
- `MSG_Error_Sheet`
- `MSG_Correct`
- `MSG_ExcelFile`
- `MSG_FichNonTr`
- `MSG_TitleGeo`
- `MSG_UpdatePass`
- `btnTransImp`
- `MSG_MovingData`
- `MSG_PassWord`
- `MSG_OpenLL`
- `MSG_TitlePassWord`
- `MSG_Title_OutPut`
- `MSG_LoadGeo`
- `MSG_CheckLL`
- `MSG_CloseLL`
- `MSG_PreparLL`
- `MSG_EnCours`
- `MSG_ProvTemp`
- `MSG_ReadExport`
- `MSG_ReadList`
- `MSG_SelectFile`
- `MSG_SelectFolder`
- `MSG_Set`
- `MSG_Continue`
- `MSG_GeoNotLoaded`
- `MSG_CloseOutPut`
- `MSG_Tempfile`
- `MSG_PathGeo`
- `MSG_Traduit`
- `MSG_Translating`
- `MSG_NotExists`
- `MSG_DefaultButton`
- `MSG_UpdatePassQ`

## Dead rows: designer worksheet labels (t_tradrange)

- `RNG_LabLangDesigner`

## Dead rows: designer worksheet buttons (t_tradshape)

- `SHP_Annuler`
- `SHP_Generer`
- `SHP_OpenLL`
- `SHP_Reset`
- `SHP_TitreMain`
- `SHP_Trad`
- `SHP_TradDico`
- `SHP_TradGeo`

## Dead rows: designer dropdowns (t_traddrop)

- `DROPEPIWEEK`

## Cells left blank

| table | tag | language |
| --- | --- | --- |
| t_tradllmsg | `MSG_Enter` | ARA |
| t_tradllmsg | `MSG_PrintSheet` | ARA |
| t_tradllmsg | `MSG_RowHeight` | ARA |

## Rows that read the same in all five languages

_Nothing._

## Rows for a ribbon control that never asks for its label

The control is in the ribbon XML, but it carries a plain `label=` or no label at all rather than `getLabel=`. The row is never read and what the user sees stays the same in every language. Either drop the row or give the control a getLabel callback.

- `customGroupAdvanced`
- `customGroupAna`

## Tags carried twice in one table

_Nothing._

## Controls left with their default name

A control still called `Label1` cannot be translated: no table row will ever be keyed on that name. Rename it on the form first.

- `OptionButton1` on F_EpiWeek
- `OptionButton2` on F_EpiWeek
- `OptionButton3` on F_EpiWeek
- `Label1` on F_ExportMig

## What was read

- VBA source: `src/classes` and `src/modules`, 19 files naming a tag.
- Forms: `.mock/forms/designer`, 11 forms naming a tag.
- Ribbons: `ribbons/_ribbontemplate_main/ribbon.xml`, `ribbons/_ribbontemplate_dev/ribbon.xml`, `ribbons/designer/ribbon.xml`, `ribbons/designer_mock/ribbon.xml`.
- Designer binary: `src/bin/designer/designer.xlsb`, 148 defined names, 5 shapes.

Left out on purpose: `src/tests`, `src/classes/stale`, and the setup and master-setup folders. The setup workbooks carry their own translation table and are not in this workbook.

