VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrossTableServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Facade wiring cross table helper instances"

'@Folder("Implement")
'@ModuleDescription("Entry point providing access to cross table helper classes")

Option Explicit

Public Function BuildLayoutPlanner(ByVal specs As ITablesSpecs, ByVal targetSheet As Worksheet) As CrossTableLayoutPlanner
    Set BuildLayoutPlanner = CrossTableLayoutPlanner.Create(specs, targetSheet)
End Function

Public Function BuildRangeRegistrar(ByVal targetSheet As Worksheet) As CrossTableRangeRegistrar
    Set BuildRangeRegistrar = CrossTableRangeRegistrar.Create(targetSheet)
End Function

Public Function BuildFormulaBuilder(ByVal evaluationSheet As Worksheet) As CrossTableFormulaBuilder
    Set BuildFormulaBuilder = CrossTableFormulaBuilder.Create(evaluationSheet)
End Function

Public Function NewConditionSet() As CrossTableConditionSet
    Dim setObject As CrossTableConditionSet
    Set setObject = New CrossTableConditionSet
    setObject.Initialise
    Set NewConditionSet = setObject
End Function

Public Function BuildWriteBuffer(ByVal startRow As Long, ByVal startColumn As Long) As CrossTableWriteBuffer
    Dim buffer As CrossTableWriteBuffer
    Set buffer = New CrossTableWriteBuffer
    buffer.Initialise startRow, startColumn
    Set BuildWriteBuffer = buffer
End Function

End Class
