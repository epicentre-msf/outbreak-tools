VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HiddenNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Manages hidden worksheet name/value storage with metadata caching"

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Manages hidden worksheet name/value storage with metadata caching")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing

Option Explicit

Implements IHiddenNames

'@class HiddenNames
'@description
'Manages hidden worksheet/workbook-level name-value storage with typed
'serialization and metadata caching. Each name carries a type annotation
'(String/Boolean/Long/Variant) and a last-updated timestamp stored in
'the Name.Comment field. Supports cross-scope import/export between
'worksheets and workbooks.
'@depends ProjectError, BetterArray
'@version 1.0 (2026-02-09)


Private Type THiddenNames
    hostSh As Worksheet
    hostWb As Workbook
    metadata As BetterArray
End Type

Private this As THiddenNames

Private Const CLASSNAME As String = "HiddenNames"
Private Const COMMENT_PREFIX As String = "HiddenNames:"
Private Const COMMENT_SEPARATOR As String = ";"
Private Const COMMENT_TYPE_PREFIX As String = "type="
Private Const COMMENT_UPDATED_PREFIX As String = "updated="
Private Const TIMESTAMP_FORMAT As String = "yyyy-mm-dd\THH:NN:SS"

Private Const META_FIELD_NAME As String = "name"
Private Const META_FIELD_TYPE As String = "type"
Private Const META_FIELD_UPDATED As String = "updated"
Private Const META_FIELD_DEFINITION As String = "definition"


'@section Metadata Helpers
'===============================================================================

'@label:metadata-field-index
'@sub-title Ordinal positions within a metadata record array
Private Enum MetadataFieldIndex
    MetaFieldName = 0
    MetaFieldType = 1
    MetaFieldUpdated = 2
    MetaFieldDefinition = 3
End Enum

'@label:metadata-store
'@sub-title Lazy-initialised BetterArray holding all metadata records
'@return BetterArray. The metadata cache.
Private Function MetadataStore() As BetterArray
    If this.metadata Is Nothing Then
        Set this.metadata = New BetterArray
        this.metadata.LowerBound = 0
    End If
    Set MetadataStore = this.metadata
End Function

'@label:build-metadata-record
'@sub-title Construct a metadata record array
Private Function BuildMetadataRecord(ByVal nameId As String, _
                                     ByVal valueType As Byte, _
                                     ByVal lastUpdated As Date, _
                                     ByVal definition As Name) As Variant
    Dim record(0 To 3) As Variant

    record(MetaFieldName) = nameId
    record(MetaFieldType) = valueType
    record(MetaFieldUpdated) = lastUpdated
    Set record(MetaFieldDefinition) = definition

    BuildMetadataRecord = record
End Function

'@label:metadata-name
'@sub-title Read the name field from a metadata record
Private Function MetadataName(ByRef record As Variant) As String
    MetadataName = CStr(record(MetaFieldName))
End Function

'@label:metadata-type
'@sub-title Read the type field from a metadata record
Private Function MetadataType(ByRef record As Variant) As Byte
    MetadataType = CLng(record(MetaFieldType))
End Function

'@label:metadata-updated
'@sub-title Read the last-updated field from a metadata record
Private Function MetadataUpdated(ByRef record As Variant) As Date
    MetadataUpdated = record(MetaFieldUpdated)
End Function

'@label:metadata-definition
'@sub-title Read the Name definition from a metadata record
Private Function MetadataDefinition(ByRef record As Variant) As Name
    On Error Resume Next
        Set MetadataDefinition = record(MetaFieldDefinition)
    On Error GoTo 0
End Function

'@label:metadata-set-definition
'@sub-title Write the Name definition into a metadata record
Private Sub MetadataSetDefinition(ByRef record As Variant, ByVal definition As Name)
    Set record(MetaFieldDefinition) = definition
End Sub

'@label:metadata-set-updated
'@sub-title Write the last-updated timestamp into a metadata record
Private Sub MetadataSetUpdated(ByRef record As Variant, ByVal lastUpdated As Date)
    record(MetaFieldUpdated) = lastUpdated
End Sub

'@label:metadata-set-type
'@sub-title Write the value type into a metadata record
Private Sub MetadataSetType(ByRef record As Variant, ByVal valueType As Byte)
    record(MetaFieldType) = valueType
End Sub

'@label:metadata-index
'@sub-title Find the index of a metadata record by name
'@return Long. The index, or -1 when not found.
Private Function MetadataIndex(ByVal nameId As String) As Long
    Dim store As BetterArray
    Dim idx As Long
    Dim record As Variant

    MetadataIndex = -1
    If this.metadata Is Nothing Then Exit Function

    Set store = this.metadata
    If store.Length = 0 Then Exit Function

    For idx = store.LowerBound To store.UpperBound
        record = store.Item(idx)
        If StrComp(MetadataName(record), nameId, vbTextCompare) = 0 Then
            MetadataIndex = idx
            Exit Function
        End If
    Next idx
End Function

'@label:store-metadata-record
'@sub-title Insert or replace a metadata record in the cache
Private Sub StoreMetadataRecord(ByVal nameId As String, ByRef record As Variant)
    Dim store As BetterArray
    Dim idx As Long

    idx = MetadataIndex(nameId)
    Set store = MetadataStore()
    If idx = -1 Then
        store.Push record
    Else
        store.Item(idx) = record
    End If
End Sub


'@section Factory
'===============================================================================

'@label:create
'@sub-title Create a new HiddenNames instance
'@details
'Validates the input and creates a fully initialised manager bound to
'the supplied worksheet or workbook.
'@param targetObj Object. A Worksheet or Workbook hosting the hidden names.
'@return IHiddenNames. The configured instance.
'@throws ProjectError.ObjectNotInitialized When targetObj is Nothing.
'@throws ProjectError.InvalidArgument When targetObj is not a Worksheet or Workbook.
'@export
Public Function Create(ByVal targetObj As Object) As IHiddenNames
    Dim manager As HiddenNames

    ValidateInput targetObj

    Set manager = New HiddenNames
    manager.Initialise targetObj
    Set Create = manager.Self
End Function

'@label:self
'@prop-title Self-reference returning the interface pointer
'@return IHiddenNames. This instance.
Public Property Get Self() As IHiddenNames
    Set Self = Me
End Property


'@section Initialisation
'===============================================================================

'@label:initialise
'@sub-title Bind the manager to a worksheet or workbook and load cache
Public Sub Initialise(ByVal targetObj As Object)
    Set this.hostSh = Nothing
    Set this.hostWb = Nothing

    Select Case TypeName(targetObj)
        Case "Worksheet"
            Set this.hostSh = targetObj
        Case "Workbook"
            Set this.hostWb = targetObj
    End Select

    Set this.metadata = Nothing
    LoadMetadataCache
End Sub

'@label:load-metadata-cache
'@sub-title Scan existing hidden names and populate the metadata cache
Private Sub LoadMetadataCache()
    Dim definition As Name
    Dim valueType As Byte
    Dim lastUpdated As Date
    Dim record As Variant
    Dim store As BetterArray
    Dim nameId As String

    Dim container As Object

    Set container = NameOwner()
    If container Is Nothing Then Exit Sub

    Set store = MetadataStore()

    For Each definition In container.Names
        If ShouldTrack(definition) Then
            nameId = ExtractSimpleName(definition.Name)
            valueType = NameTypeFromComment(definition.Comment)
            lastUpdated = UpdatedFromComment(definition.Comment)
            record = BuildMetadataRecord(nameId, valueType, lastUpdated, definition)
            store.Push record
        End If
    Next definition
End Sub


'@section Interface Helpers
'===============================================================================

'@label:validate-input
'@sub-title Validate the factory input argument
Private Sub ValidateInput(ByVal targetObj As Object)

    If targetObj Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "worksheet/workbook required to create hidden names" & CLASSNAME
    End If


    If( TypeName(targetObj) <> "Worksheet") And (TypeName(targetObj) <> "Workbook")  Then
        ThrowError ProjectError.InvalidArgument, "Hidden names needs an object of type worksheet or workbook"
    End If
End Sub

'@label:sheet
'@prop-title Worksheet backing this manager
'@export
Private Property Get SheetInternal() As Worksheet
    Set SheetInternal = this.hostSh
End Property

'@label:name-owner
'@sub-title Resolve the host object owning the names collection
Private Function NameOwner() As Object
    If Not this.hostSh Is Nothing Then
        Set NameOwner = this.hostSh
    ElseIf Not this.hostWb Is Nothing Then
        Set NameOwner = this.hostWb
    End If
End Function

'@label:scope-label
'@sub-title Human-readable label for the current scope
Private Function ScopeLabel() As String
    If Not this.hostSh Is Nothing Then
        ScopeLabel = "sheet '" & this.hostSh.Name & "'"
    ElseIf Not this.hostWb Is Nothing Then
        ScopeLabel = "workbook '" & this.hostWb.Name & "'"
    Else
        ScopeLabel = "the configured scope"
    End If
End Function

'@label:sanitize-name-id
'@sub-title Convert a raw identifier to a valid Excel Name
'@details
'Excel defined names only allow letters, digits, underscores, and periods.
'The first character must be a letter or underscore.
'This function replaces any disallowed character (spaces, hyphens, etc.)
'with an underscore and prepends one when the name starts with a digit
'or period.  Called once at each public API entry point so that all
'internal methods receive pre-sanitized identifiers.
'@param nameId String. The raw name identifier.
'@return String. An Excel-safe name identifier.
Private Function SanitizeNameId(ByVal nameId As String) As String
    Dim raw As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim result As String

    raw = Trim$(nameId)
    If LenB(raw) = 0 Then
        SanitizeNameId = vbNullString
        Exit Function
    End If

    result = vbNullString
    For i = 1 To Len(raw)
        code = AscW(Mid$(raw, i, 1))
        ' Allow A-Z (65-90), a-z (97-122), 0-9 (48-57), underscore (95), period (46)
        If (code >= 65 And code <= 90) Or _
           (code >= 97 And code <= 122) Or _
           (code >= 48 And code <= 57) Or _
           code = 95 Or code = 46 Then
            result = result & Mid$(raw, i, 1)
        Else
            result = result & "_"
        End If
    Next i

    ' First character must be a letter or underscore (not digit or period)
    If Len(result) > 0 Then
        code = AscW(Left$(result, 1))
        If Not ((code >= 65 And code <= 90) Or _
                (code >= 97 And code <= 122) Or _
                code = 95) Then
            result = "_" & result
        End If
    End If

    SanitizeNameId = result
End Function

'@label:metadata-key
'@sub-title Normalise a name identifier for case-insensitive lookup
'@details Expects a pre-sanitized nameId (callers must call SanitizeNameId first).
Private Function MetadataKey(ByVal nameId As String) As String
    MetadataKey = LCase$(nameId)
End Function

'@label:ensure-metadata-record
'@sub-title Ensure a metadata record exists for the given name
'@details Expects a pre-sanitized nameId (callers must call SanitizeNameId first).
'@return Long. The record index, or -1 when the name is not defined.
Private Function EnsureMetadataRecord(ByVal nameId As String) As Long
    Dim idx As Long
    Dim store As BetterArray
    Dim record As Variant
    Dim definition As Name
    Dim valueType As Byte
    Dim lastUpdated As Date

    idx = MetadataIndex(nameId)
    Set store = MetadataStore()

    If idx = -1 Then
        Set definition = FindDefinition(nameId)
        If definition Is Nothing Then
            EnsureMetadataRecord = -1
            Exit Function
        End If

        valueType = NameTypeFromComment(definition.Comment)
        lastUpdated = UpdatedFromComment(definition.Comment)
        record = BuildMetadataRecord(nameId, valueType, lastUpdated, definition)
        store.Push record
        EnsureMetadataRecord = store.UpperBound
        Exit Function
    End If

    record = store.Item(idx)
    If MetadataDefinition(record) Is Nothing Then
        Set definition = FindDefinition(nameId)
        If Not definition Is Nothing Then
            MetadataSetDefinition record, definition
            store.Item(idx) = record
        End If
    End If

    EnsureMetadataRecord = idx
End Function

'@label:metadata-record
'@sub-title Retrieve a metadata record by name
Private Function MetadataRecord(ByVal nameId As String) As Variant
    Dim idx As Long
    Dim store As BetterArray

    idx = EnsureMetadataRecord(nameId)
    If idx = -1 Then Exit Function

    Set store = MetadataStore()
    MetadataRecord = store.Item(idx)
End Function

'@label:remove-metadata-record
'@sub-title Remove a metadata record from the cache
Private Sub RemoveMetadataRecord(ByVal nameId As String)
    Dim idx As Long
    Dim store As BetterArray
    Dim buffer As BetterArray
    Dim i As Long

    idx = MetadataIndex(nameId)
    If idx = -1 Then Exit Sub

    Set store = MetadataStore()
    Set buffer = New BetterArray
    buffer.LowerBound = store.LowerBound

    For i = store.LowerBound To store.UpperBound
        If i <> idx Then
            buffer.Push store.Item(i)
        End If
    Next i

    Set this.metadata = buffer
End Sub


'@section Lookup Helpers
'===============================================================================

'@label:should-track
'@sub-title Determine whether a Name definition belongs to this scope
Private Function ShouldTrack(ByVal definition As Name) As Boolean
    Dim parentSh As Worksheet
    Dim refersTo As String
    Dim container As Object
    Dim parentType As String

    If definition Is Nothing Then Exit Function
    If definition.Visible Then Exit Function

    Set container = NameOwner()
    If container Is Nothing Then Exit Function

    parentType = TypeName(definition.Parent)

    If TypeName(container) = "Worksheet" Then
        If StrComp(parentType, "Worksheet", vbTextCompare) = 0 Then
            Set parentSh = definition.Parent
            ShouldTrack = (StrComp(parentSh.Name, container.Name, vbTextCompare) = 0)
        Else
            On Error Resume Next
                refersTo = definition.RefersTo
            On Error GoTo 0
            If LenB(refersTo) = 0 Then Exit Function
            ShouldTrack = (InStr(1, refersTo, "'" & container.Name & "'!", vbTextCompare) > 0)
        End If
    Else
        If StrComp(parentType, "Workbook", vbTextCompare) = 0 Then
            ShouldTrack = (definition.Parent Is container)
        End If
    End If
End Function

'@label:find-definition
'@sub-title Locate a Name definition by identifier
'@details Expects a pre-sanitized nameId (callers must call SanitizeNameId first).
Private Function FindDefinition(ByVal nameId As String) As Name
    Dim definition As Name
    Dim container As Object

    Set container = NameOwner()
    If container Is Nothing Then Exit Function

    If LenB(nameId) = 0 Then Exit Function

    For Each definition In container.Names
        If ShouldTrack(definition) Then
            If CompareName(definition, nameId) Then
                Set FindDefinition = definition
                Exit Function
            End If
        End If
    Next definition
End Function

'@label:compare-name
'@sub-title Case-insensitive comparison of a Name definition to an expected identifier
Private Function CompareName(ByVal definition As Name, ByVal expected As String) As Boolean
    CompareName = (StrComp(ExtractSimpleName(definition.Name), expected, vbTextCompare) = 0)
End Function

'@label:extract-simple-name
'@sub-title Strip sheet qualification from a Name string
Private Function ExtractSimpleName(ByVal qualifiedName As String) As String
    Dim exclPos As Long

    exclPos = InStr(qualifiedName, "!")
    If exclPos = 0 Then
        ExtractSimpleName = qualifiedName
    Else
        ExtractSimpleName = Mid$(qualifiedName, exclPos + 1)
    End If
End Function


'@section Serialization
'===============================================================================

'@label:serialize-value
'@sub-title Convert a typed value to its RefersTo formula string
Private Function SerializeValue(ByVal value As Variant, _
                                ByVal valueType As Byte) As String
    Select Case valueType
        Case HiddenNameTypeString
            SerializeValue = """" & EscapeString(CStr(value)) & """"
        Case HiddenNameTypeBoolean
            If CBool(value) Then
                SerializeValue = "TRUE"
            Else
                SerializeValue = "FALSE"
            End If
        Case HiddenNameTypeLong
            SerializeValue = CStr(CLng(value))
        Case Else
            SerializeValue = SerializeVariant(value)
    End Select
End Function

'@label:serialize-variant
'@sub-title Serialize a Variant value by inspecting its runtime type
Private Function SerializeVariant(ByVal value As Variant) As String
    Select Case VarType(value)
        Case vbString
            SerializeVariant = """" & EscapeString(value) & """"
        Case vbBoolean
            If value Then
                SerializeVariant = "TRUE"
            Else
                SerializeVariant = "FALSE"
            End If
        Case vbLong, vbInteger, vbByte
            SerializeVariant = CStr(CLng(value))
        Case vbDouble, vbSingle, vbCurrency
            SerializeVariant = CStr(value)
        Case vbDate
            SerializeVariant = """" & Format$(value, "yyyy-mm-dd\THH:NN:SS") & """"
        Case Else
            SerializeVariant = """" & EscapeString(CStr(value)) & """"
    End Select
End Function

'@label:escape-string
'@sub-title Escape double-quotes for formula embedding
Private Function EscapeString(ByVal value As String) As String
    EscapeString = Replace(value, """", """""")
End Function

'@label:apply-comment
'@sub-title Write the type and timestamp comment onto a Name definition
Private Sub ApplyComment(ByVal definition As Name, _
                         ByVal valueType As Byte, _
                         ByVal updatedOn As Date)
    On Error Resume Next
        definition.Comment = BuildComment(valueType, updatedOn)
    On Error GoTo 0
End Sub

'@label:build-comment
'@sub-title Assemble the structured comment string
Private Function BuildComment(ByVal valueType As Byte, ByVal updatedOn As Date) As String
    Dim comment As String

    comment = COMMENT_PREFIX & COMMENT_TYPE_PREFIX & ValueTypeName(valueType)
    comment = comment & COMMENT_SEPARATOR & COMMENT_UPDATED_PREFIX & Format$(updatedOn, TIMESTAMP_FORMAT)
    BuildComment = comment
End Function

'@label:name-type-from-comment
'@sub-title Parse the value type from a Name comment
Private Function NameTypeFromComment(ByVal comment As String) As Byte
    Dim typeValue As String

    typeValue = ExtractCommentValue(comment, COMMENT_TYPE_PREFIX)
    If LenB(typeValue) = 0 Then
        NameTypeFromComment = HiddenNameTypeVariant
    Else
        NameTypeFromComment = StringToValueType(typeValue)
    End If
End Function

'@label:updated-from-comment
'@sub-title Parse the last-updated timestamp from a Name comment
Private Function UpdatedFromComment(ByVal comment As String) As Date
    Dim stamp As String

    stamp = ExtractCommentValue(comment, COMMENT_UPDATED_PREFIX)
    If LenB(stamp) = 0 Then Exit Function

    On Error Resume Next
        UpdatedFromComment = CDate(stamp)
    On Error GoTo 0
End Function

'@label:extract-comment-value
'@sub-title Extract a key-value token from a structured comment
Private Function ExtractCommentValue(ByVal comment As String, ByVal prefix As String) As String
    Dim lowerComment As String
    Dim lowerPrefix As String
    Dim startPos As Long
    Dim endPos As Long

    If LenB(comment) = 0 Then Exit Function

    lowerComment = LCase$(comment)
    lowerPrefix = LCase$(prefix)

    startPos = InStr(lowerComment, lowerPrefix)
    If startPos = 0 Then Exit Function

    startPos = startPos + Len(prefix)
    endPos = InStr(startPos, comment, COMMENT_SEPARATOR)

    If endPos = 0 Then
        ExtractCommentValue = Mid$(comment, startPos)
    Else
        ExtractCommentValue = Mid$(comment, startPos, endPos - startPos)
    End If
    ExtractCommentValue = Trim$(ExtractCommentValue)
End Function

'@label:value-type-name
'@sub-title Convert a HiddenNameValueType to its string representation
Private Function ValueTypeName(ByVal valueType As Byte) As String
    Select Case valueType
        Case HiddenNameTypeString: ValueTypeName = "string"
        Case HiddenNameTypeBoolean: ValueTypeName = "boolean"
        Case HiddenNameTypeLong: ValueTypeName = "long"
        Case Else: ValueTypeName = "variant"
    End Select
End Function

'@label:string-to-value-type
'@sub-title Convert a string representation back to HiddenNameValueType
Private Function StringToValueType(ByVal value As String) As Byte
    Select Case LCase$(value)
        Case "string": StringToValueType = HiddenNameTypeString
        Case "boolean": StringToValueType = HiddenNameTypeBoolean
        Case "long": StringToValueType = HiddenNameTypeLong
        Case Else: StringToValueType = HiddenNameTypeVariant
    End Select
End Function


'@section Value Coercion
'===============================================================================

'@label:coerce-value
'@sub-title Cast a raw value to the specified HiddenNameValueType
Private Function CoerceValue(ByVal value As Variant, _
                             ByVal valueType As Byte) As Variant
    Select Case valueType
        Case HiddenNameTypeString
            CoerceValue = CStr(value)
        Case HiddenNameTypeBoolean
            CoerceValue = CBool(value)
        Case HiddenNameTypeLong
            CoerceValue = CLng(value)
        Case Else
            CoerceValue = value
    End Select
End Function


'@section Core Operations
'===============================================================================

'@label:ensure-definition
'@sub-title Create a hidden Name definition if it does not exist
'@details Expects a pre-sanitized nameId (callers must call SanitizeNameId first).
Private Function EnsureDefinition(ByVal nameId As String, _
                                  ByVal initialValue As Variant, _
                                  ByVal valueType As Byte) As Name
    Dim definition As Name
    Dim nowStamp As Date

    If LenB(nameId) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Hidden name identifier cannot be empty"
    End If

    Dim container As Object

    Set container = NameOwner()
    If container Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "HiddenNames requires a target worksheet or workbook."
    End If

    Set definition = FindDefinition(nameId)
    If definition Is Nothing Then
        nowStamp = Now
        Set definition = container.Names.Add(Name:=nameId, _
                                             RefersTo:="=" & SerializeValue(initialValue, valueType), _
                                             Visible:=False)
        ApplyComment definition, valueType, nowStamp
    End If

    Set EnsureDefinition = definition
End Function

'@label:update-definition-value
'@sub-title Write a new value into an existing Name definition
Private Sub UpdateDefinitionValue(ByVal definition As Name, _
                                  ByVal value As Variant, _
                                  ByVal valueType As Byte, _
                                  Optional ByVal updatedOn As Date = 0)
    Dim timestamp As Date

    If updatedOn = 0 Then
        timestamp = Now
    Else
        timestamp = updatedOn
    End If

    definition.RefersTo = "=" & SerializeValue(value, valueType)
    definition.Visible = False
    ApplyComment definition, valueType, timestamp
End Sub

'@label:definition-stored-value
'@sub-title Extract the raw stored value from a Name definition
Private Function DefinitionStoredValue(ByVal definition As Name, _
                                       ByVal valueType As Byte) As Variant
    Dim raw As Variant
    Dim textValue As String

    On Error Resume Next
        raw = definition.Value
    On Error GoTo 0

    If VarType(raw) = vbString Then
        textValue = Trim$(CStr(raw))
        If LenB(textValue) > 0 Then
            If Left$(textValue, 1) = "=" Then
                textValue = Mid$(textValue, 2)
            End If
        End If

        If valueType = HiddenNameTypeString Then
            If Len(textValue) >= 2 Then
                If Left$(textValue, 1) = """" And Right$(textValue, 1) = """" Then
                    textValue = Mid$(textValue, 2, Len(textValue) - 2)
                End If
            End If
            textValue = Replace(textValue, """""", """")
        End If

        DefinitionStoredValue = textValue
    Else
        DefinitionStoredValue = raw
    End If
End Function

'@label:remove-definition
'@sub-title Delete a Name definition from the host scope
Private Sub RemoveDefinition(ByVal nameId As String)
    Dim definition As Name

    Set definition = FindDefinition(nameId)
    If definition Is Nothing Then Exit Sub

    On Error Resume Next
        definition.Delete
    On Error GoTo 0
End Sub

'@label:ensure-ready
'@sub-title Guard that the manager is properly initialised
Private Sub EnsureReady()
    If NameOwner() Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                  "HiddenNames requires a worksheet or workbook. Use Create before calling members."
    End If
End Sub


'@section Public API
'===============================================================================

'@label:has-name
'@sub-title Check whether a hidden name exists
'@export
Private Function HasName(ByVal nameId As String) As Boolean
    HasName = (EnsureMetadataRecord(SanitizeNameId(nameId)) <> -1)
End Function

'@label:ensure-name
'@sub-title Ensure a hidden name exists with the provided default
'@export
Private Function Ensure(ByVal nameId As String, _
                        ByVal initialValue As Variant, _
                        Optional ByVal valueType As Byte = HiddenNameTypeVariant) As Name
    Dim sanitized As String
    Dim definition As Name
    Dim record As Variant
    Dim lastUpdated As Date

    EnsureReady
    sanitized = SanitizeNameId(nameId)
    Set definition = EnsureDefinition(sanitized, initialValue, valueType)
    lastUpdated = Now
    record = BuildMetadataRecord(sanitized, valueType, lastUpdated, definition)
    StoreMetadataRecord sanitized, record

    Set Ensure = definition
End Function

'@label:value
'@sub-title Retrieve a hidden value without coercion
'@export
Private Function ValueOf(ByVal nameId As String, _
                         Optional ByVal defaultValue As Variant) As Variant
    Dim sanitized As String
    Dim idx As Long
    Dim store As BetterArray
    Dim record As Variant
    Dim definition As Name
    Dim raw As Variant

    EnsureReady
    sanitized = SanitizeNameId(nameId)
    idx = EnsureMetadataRecord(sanitized)

    If idx = -1 Then
        If IsMissing(defaultValue) Then
            ThrowError ProjectError.ElementNotFound, _
                      "Hidden name '" & nameId & "' is not defined on " & ScopeLabel()
        Else
            ValueOf = defaultValue
        End If
        Exit Function
    End If

    Set store = MetadataStore()
    record = store.Item(idx)
    Set definition = MetadataDefinition(record)
    raw = DefinitionStoredValue(definition, MetadataType(record))
    ValueOf = CoerceValue(raw, MetadataType(record))
End Function

'@label:set-value
'@sub-title Update the stored value of an existing hidden name
'@export
Private Sub Update(ByVal nameId As String, ByVal value As Variant)
    Dim sanitized As String
    Dim idx As Long
    Dim store As BetterArray
    Dim record As Variant
    Dim definition As Name
    Dim nowStamp As Date

    EnsureReady
    sanitized = SanitizeNameId(nameId)
    idx = EnsureMetadataRecord(sanitized)
    If idx = -1 Then
        ThrowError ProjectError.ElementNotFound, _
                  "Cannot update hidden name '" & nameId & "' because it does not exist."
    End If

    Set store = MetadataStore()
    record = store.Item(idx)
    Set definition = MetadataDefinition(record)
    If definition Is Nothing Then
        Set definition = EnsureDefinition(MetadataName(record), value, MetadataType(record))
    End If

    UpdateDefinitionValue definition, value, MetadataType(record)
    nowStamp = Now
    MetadataSetDefinition record, definition
    MetadataSetUpdated record, nowStamp
    store.Item(idx) = record
End Sub

'@label:value-as-string
'@sub-title Retrieve the hidden value coerced to String
'@export
Private Function ValueAsString(ByVal nameId As String, _
                               Optional ByVal defaultValue As String = vbNullString) As String
    Dim value As Variant

    value = ValueOf(nameId, defaultValue)
    ValueAsString = CStr(value)
End Function

'@label:value-as-boolean
'@sub-title Retrieve the hidden value coerced to Boolean
'@export
Private Function ValueAsBoolean(ByVal nameId As String, _
                                Optional ByVal defaultValue As Boolean = False) As Boolean
    Dim value As Variant

    value = ValueOf(nameId, defaultValue)
    ValueAsBoolean = CBool(value)
End Function

'@label:value-as-long
'@sub-title Retrieve the hidden value coerced to Long
'@export
Private Function ValueAsLong(ByVal nameId As String, _
                             Optional ByVal defaultValue As Long = 0) As Long
    Dim value As Variant

    value = ValueOf(nameId, defaultValue)
    ValueAsLong = CLng(value)
End Function

'@label:remove-name
'@sub-title Remove a hidden name definition
'@export
Private Sub Remove(ByVal nameId As String)
    Dim sanitized As String

    EnsureReady
    sanitized = SanitizeNameId(nameId)
    RemoveDefinition sanitized
    RemoveMetadataRecord sanitized
End Sub

'@label:list-names
'@sub-title Enumerate recorded hidden names filtered by optional prefix
'@export
Private Function Enumerate(Optional ByVal prefix As String = vbNullString) As BetterArray
    Dim result As BetterArray
    Dim store As BetterArray
    Dim idx As Long
    Dim record As Variant
    Dim prefixKey As String
    Dim includeEntry As Boolean
    Dim nameId As String
    Dim valueType As Byte
    Dim lastUpdated As Date

    EnsureReady
    Set result = New BetterArray
    result.LowerBound = 0
    prefixKey = LCase$(SanitizeNameId(prefix))

    If this.metadata Is Nothing Then
        Set Enumerate = result
        Exit Function
    End If

    Set store = MetadataStore()
    If store.Length = 0 Then
        Set Enumerate = result
        Exit Function
    End If

    For idx = store.LowerBound To store.UpperBound
        record = store.Item(idx)
        nameId = MetadataName(record)
        includeEntry = (LenB(prefixKey) = 0) Or _
                       (Left$(MetadataKey(nameId), Len(prefixKey)) = prefixKey)

        If includeEntry Then
            valueType = MetadataType(record)
            lastUpdated = MetadataUpdated(record)
            result.Push Array(nameId, valueType, lastUpdated)
        End If
    Next idx

    Set Enumerate = result
End Function

'@label:export-names
'@sub-title Export all tracked hidden names to another worksheet
'@export
Private Sub ExportAll(ByVal targetSh As Worksheet, _
                      Optional ByVal prefix As String = vbNullString)
    Dim store As BetterArray
    Dim idx As Long
    Dim record As Variant
    Dim prefixKey As String
    Dim includeEntry As Boolean
    Dim value As Variant
    Dim nameId As String
    Dim valueType As Byte

    EnsureReady
    If targetSh Is Nothing Then
        ThrowError ProjectError.InvalidArgument, "Target sheet is required for export."
    End If

    If this.metadata Is Nothing Then Exit Sub

    prefixKey = LCase$(SanitizeNameId(prefix))
    Set store = MetadataStore()

    For idx = store.LowerBound To store.UpperBound
        record = store.Item(idx)
        nameId = MetadataName(record)
        includeEntry = (LenB(prefixKey) = 0) Or _
                       (Left$(MetadataKey(nameId), Len(prefixKey)) = prefixKey)

        If includeEntry Then
            valueType = MetadataType(record)
            value = ValueOf(nameId)
            PersistOnWorksheet targetSh, nameId, value, valueType
        End If
    Next idx
End Sub

'@label:export-names-to-workbook
'@sub-title Export tracked names to another workbook scope
'@export
Private Sub ExportAllToWorkbook(ByVal targetWb As Workbook, _
                                Optional ByVal prefix As String = vbNullString)
    Dim store As BetterArray
    Dim idx As Long
    Dim record As Variant
    Dim prefixKey As String
    Dim includeEntry As Boolean
    Dim nameId As String
    Dim valueType As Byte
    Dim value As Variant
    Dim destination As IHiddenNames

    EnsureReady
    If targetWb Is Nothing Then
        ThrowError ProjectError.InvalidArgument, "Target workbook is required for export."
    End If

    If this.metadata Is Nothing Then Exit Sub

    prefixKey = LCase$(SanitizeNameId(prefix))
    Set store = MetadataStore()
    If store.Length = 0 Then Exit Sub

    Set destination = HiddenNames.Create(targetWb)

    For idx = store.LowerBound To store.UpperBound
        record = store.Item(idx)
        nameId = MetadataName(record)
        includeEntry = (LenB(prefixKey) = 0) Or _
                       (Left$(MetadataKey(nameId), Len(prefixKey)) = prefixKey)

        If includeEntry Then
            valueType = MetadataType(record)
            value = ValueOf(nameId)
            destination.EnsureName nameId, value, valueType
            destination.SetValue nameId, value
        End If
    Next idx
End Sub

'@label:import-names
'@sub-title Import hidden names from another worksheet
'@export
Private Sub ImportAll(ByVal sourceSh As Worksheet, _
                      Optional ByVal prefix As String = vbNullString, _
                      Optional ByVal overwriteExisting As Boolean = True)
    Dim sourceManager As IHiddenNames
    Dim names As BetterArray
    Dim idx As Long
    Dim record As Variant
    Dim nameId As String
    Dim entryType As Byte
    Dim lastUpdated As Date
    Dim value As Variant

    EnsureReady
    If sourceSh Is Nothing Then
        ThrowError ProjectError.InvalidArgument, "Source sheet is required for import."
    End If

    Set sourceManager = HiddenNames.Create(sourceSh)
    Set names = sourceManager.ListNames(prefix)
    If names Is Nothing Then Exit Sub
    If names.Length = 0 Then Exit Sub

    For idx = names.LowerBound To names.UpperBound
        record = names.Item(idx)
        nameId = CStr(record(0))
        entryType = record(1)
        lastUpdated = 0
        If UBound(record) >= 2 Then lastUpdated = record(2)
        value = sourceManager.Value(nameId)

        Call Ensure(nameId, value, entryType)

        If overwriteExisting Or Not HasName(nameId) Then
            UpdateStoredValue nameId, value, entryType, lastUpdated
        End If
    Next idx
End Sub

'@label:import-names-from-workbook
'@sub-title Import hidden names from another workbook scope
'@export
Private Sub ImportAllFromWorkbook(ByVal sourceWb As Workbook, _
                                  Optional ByVal prefix As String = vbNullString, _
                                  Optional ByVal overwriteExisting As Boolean = True)
    Dim sourceManager As IHiddenNames
    Dim names As BetterArray
    Dim idx As Long
    Dim record As Variant
    Dim nameId As String
    Dim entryType As Byte
    Dim lastUpdated As Date
    Dim value As Variant

    EnsureReady
    If sourceWb Is Nothing Then
        ThrowError ProjectError.InvalidArgument, "Source workbook is required for import."
    End If

    Set sourceManager = HiddenNames.Create(sourceWb)
    Set names = sourceManager.ListNames(prefix)
    If names Is Nothing Then Exit Sub
    If names.Length = 0 Then Exit Sub

    For idx = names.LowerBound To names.UpperBound
        record = names.Item(idx)
        nameId = CStr(record(0))
        entryType = record(1)
        lastUpdated = 0
        If UBound(record) >= 2 Then lastUpdated = record(2)
        value = sourceManager.Value(nameId)

        Call Ensure(nameId, value, entryType)

        If overwriteExisting Or Not HasName(nameId) Then
            UpdateStoredValue nameId, value, entryType, lastUpdated
        End If
    Next idx
End Sub

'@label:set-list-object-header
'@sub-title Create a workbook-level hidden name referencing a ListObject header
'@export
Public Sub SetListObjectHeader(ByVal nameId As String, _
                               ByVal Lo As ListObject, _
                               ByVal headerName As String)
    Dim sanitizedName As String
    Dim sanitizedHeader As String
    Dim targetWorkbook As Workbook
    Dim headerColumn As ListColumn
    Dim formulaText As String

    sanitizedName = SanitizeNameId(nameId)
    sanitizedHeader = Trim$(headerName)

    If LenB(sanitizedName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Hidden name identifier cannot be empty when linking to a ListObject header."
    End If
    If Lo Is Nothing Then
        ThrowError ProjectError.InvalidArgument, "ListObject reference is required when linking a header."
    End If
    If LenB(sanitizedHeader) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Header name cannot be empty when linking a ListObject column."
    End If

    On Error Resume Next
        Set headerColumn = Lo.ListColumns(sanitizedHeader)
    On Error GoTo 0
    If headerColumn Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "ListObject column '" & sanitizedHeader & "' was not found."
    End If

    If Not this.hostWb Is Nothing Then
        Set targetWorkbook = this.hostWb
    ElseIf Not this.hostSh Is Nothing Then
        Set targetWorkbook = this.hostSh.Parent
    Else
        Set targetWorkbook = Lo.Parent.Parent
    End If

    If targetWorkbook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Unable to resolve workbook scope for SetListObjectHeader."
    End If

    RemoveWorkbookName targetWorkbook, sanitizedName
    formulaText = "=" & Lo.Name & "[" & headerColumn.Name & "]"

    targetWorkbook.Names.Add Name:=sanitizedName, _
                             RefersToR1C1:=formulaText, _
                             Visible:=False
End Sub


'@section Internal Export Helpers
'===============================================================================

'@label:persist-on-worksheet
'@sub-title Write a hidden name onto a target worksheet
'@details Expects a pre-sanitized nameId (callers must call SanitizeNameId first).
Private Sub PersistOnWorksheet(ByVal targetSh As Worksheet, _
                               ByVal nameId As String, _
                               ByVal value As Variant, _
                               ByVal valueType As Byte)
    Dim definition As Name
    Dim timestamp As Date

    timestamp = Now
    RemoveWorksheetName targetSh, nameId

    Set definition = targetSh.Names.Add(Name:=nameId, _
                                           RefersTo:="=" & SerializeValue(value, valueType), _
                                           Visible:=False)
    ApplyComment definition, valueType, timestamp
End Sub

'@label:remove-worksheet-name
'@sub-title Delete a Name definition from a specific worksheet
Private Sub RemoveWorksheetName(ByVal targetSh As Worksheet, ByVal nameId As String)
    Dim definition As Name

    For Each definition In targetSh.Names
        If CompareName(definition, nameId) Then
            On Error Resume Next
                definition.Delete
            On Error GoTo 0
            Exit For
        End If
    Next definition
End Sub

'@label:remove-workbook-name
'@sub-title Delete a Name definition from a workbook scope
'@details Expects a pre-sanitized nameId (callers must call SanitizeNameId first).
Private Sub RemoveWorkbookName(ByVal targetWb As Workbook, ByVal nameId As String)
    If targetWb Is Nothing Then Exit Sub

    On Error Resume Next
        targetWb.Names(nameId).Delete
    On Error GoTo 0
End Sub


'@label:update-stored-value
'@sub-title Write a value and update the metadata cache
Private Sub UpdateStoredValue(ByVal nameId As String, _
                              ByVal value As Variant, _
                              ByVal valueType As Byte, _
                              Optional ByVal updatedOn As Date = 0)
    Dim idx As Long
    Dim store As BetterArray
    Dim record As Variant
    Dim definition As Name
    Dim effectiveUpdated As Date

    idx = EnsureMetadataRecord(nameId)
    If idx = -1 Then
        Call Ensure(nameId, value, valueType)
        idx = EnsureMetadataRecord(nameId)
        If idx = -1 Then Exit Sub
    End If

    Set store = MetadataStore()
    record = store.Item(idx)
    Set definition = MetadataDefinition(record)
    If definition Is Nothing Then
        Set definition = EnsureDefinition(MetadataName(record), value, valueType)
    End If

    UpdateDefinitionValue definition, value, valueType, updatedOn

    If updatedOn = 0 Then
        effectiveUpdated = Now
    Else
        effectiveUpdated = updatedOn
    End If

    MetadataSetDefinition record, definition
    MetadataSetUpdated record, effectiveUpdated
    MetadataSetType record, valueType
    store.Item(idx) = record
End Sub


'@section Error Handling
'===============================================================================

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub


'@section Interface
'===============================================================================
Private Property Get IHiddenNames_Sheet() As Worksheet
    Set IHiddenNames_Sheet = SheetInternal
End Property

Private Function IHiddenNames_HasName(ByVal nameId As String) As Boolean
    IHiddenNames_HasName = HasName(nameId)
End Function

Private Function IHiddenNames_EnsureName(ByVal nameId As String, _
                                         ByVal initialValue As Variant, _
                                         Optional ByVal valueType As Byte = HiddenNameTypeVariant) As Name
    Set IHiddenNames_EnsureName = Ensure(nameId, initialValue, valueType)
End Function

Private Function IHiddenNames_Value(ByVal nameId As String, _
                                    Optional ByVal defaultValue As Variant) As Variant
    IHiddenNames_Value = ValueOf(nameId, defaultValue)
End Function

Private Sub IHiddenNames_SetValue(ByVal nameId As String, ByVal value As Variant)
    Update nameId, value
End Sub

Private Function IHiddenNames_ValueAsString(ByVal nameId As String, _
                                            Optional ByVal defaultValue As String = vbNullString) As String
    IHiddenNames_ValueAsString = ValueAsString(nameId, defaultValue)
End Function

Private Function IHiddenNames_ValueAsBoolean(ByVal nameId As String, _
                                             Optional ByVal defaultValue As Boolean = False) As Boolean
    IHiddenNames_ValueAsBoolean = ValueAsBoolean(nameId, defaultValue)
End Function

Private Function IHiddenNames_ValueAsLong(ByVal nameId As String, _
                                          Optional ByVal defaultValue As Long = 0) As Long
    IHiddenNames_ValueAsLong = ValueAsLong(nameId, defaultValue)
End Function

Private Sub IHiddenNames_RemoveName(ByVal nameId As String)
    Remove nameId
End Sub

Private Function IHiddenNames_ListNames(Optional ByVal prefix As String = vbNullString) As BetterArray
    Set IHiddenNames_ListNames = Enumerate(prefix)
End Function

Private Sub IHiddenNames_ExportNames(ByVal targetSh As Worksheet, _
                                     Optional ByVal prefix As String = vbNullString)
    ExportAll targetSh, prefix
End Sub

Private Sub IHiddenNames_ExportNamesToWorkbook(ByVal targetWb As Workbook, _
                                               Optional ByVal prefix As String = vbNullString)
    ExportAllToWorkbook targetWb, prefix
End Sub

Private Sub IHiddenNames_ImportNames(ByVal sourceSh As Worksheet, _
                                     Optional ByVal prefix As String = vbNullString, _
                                     Optional ByVal overwriteExisting As Boolean = True)
    ImportAll sourceSh, prefix, overwriteExisting
End Sub

Private Sub IHiddenNames_ImportNamesFromWorkbook(ByVal sourceWb As Workbook, _
                                                 Optional ByVal prefix As String = vbNullString, _
                                                 Optional ByVal overwriteExisting As Boolean = True)
    ImportAllFromWorkbook sourceWb, prefix, overwriteExisting
End Sub

Private Sub IHiddenNames_SetListObjectHeader(ByVal nameId As String, _
                                             ByVal listObject As ListObject, _
                                             ByVal headerName As String)
    SetListObjectHeader nameId, listObject, headerName
End Sub
