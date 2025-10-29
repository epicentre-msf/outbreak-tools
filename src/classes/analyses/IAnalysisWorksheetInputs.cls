VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisWorksheetInputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Provides worksheet-level dependencies required by the coordinator pipeline")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Return the linelist specifications used across worksheets.
Public Property Get Linelist() As ILinelistSpecs
End Property

'Return the output worksheet for the specified analysis scope.
Public Function OutputSheet(ByVal scope As AnalysisScope) As Worksheet
End Function

'Return the listobject from the source analysis sheet by name.
Public Function SpecificationList(ByVal listId As String) As ListObject
End Function

'Return the translation object for label localisation.
Public Property Get Translator() As ITranslationObject
End Property

'Return the names registry used to capture table and graph ranges.
Public Property Get NamesRegistry() As IAnaTabIds
End Property

'Return the design formatter to apply table/graph styling.
Public Property Get DesignFormatter() As ILLFormat
End Property

'Return the GoTo label provider used to localise dropdown captions.
Public Function GoToLabel(ByVal scope As String) As String
End Function
