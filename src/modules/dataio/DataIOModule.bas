Attribute VB_Name = "DataIOModule"

'@Folder("DataIO")
'@ModuleDescription("Shared DataIO utilities and entry points")

Option Explicit


' @description Check whether the linelist workbook has any entered data.
' Utility wrapper around LLImporter.HasData for use by ribbon/form code
' that needs to check data state without creating a full importer.
' @param sourceWkb Workbook. The linelist workbook to check.
' @return True when at least one HList row has user data.
Public Function LinelistHasData(ByVal sourceWkb As Workbook) As Boolean
    Dim imp As ILLImporter
    Set imp = LLImporter.Create(sourceWkb)
    LinelistHasData = imp.HasData
End Function
