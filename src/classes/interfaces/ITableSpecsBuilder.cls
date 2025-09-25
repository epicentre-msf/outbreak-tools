VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITableSpecsBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Factory contract used by TableSpecsNavigator to materialise specs")
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Create a TableSpecs instance for the supplied row.
Public Function CreateSpec(ByVal headerRange As Range, _
                           ByVal tableRange As Range, _
                           ByVal linelistSpecs As ILinelistSpecs) As ITablesSpecs
End Function
