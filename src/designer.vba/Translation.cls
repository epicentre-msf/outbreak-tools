VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Translation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private mEN As String
Private mFR As String
Private mSP As String

Property Get EN() As String
    EN = mEN
End Property

Property Let EN(value As String)
    mEN = value
End Property

Property Get FR() As String
    FR = mFR
End Property

Property Let FR(value As String)
    mFR = value
End Property

Property Get SP() As String
    SP = mSP
End Property

Property Let SP(value As String)
    mSP = value
End Property

