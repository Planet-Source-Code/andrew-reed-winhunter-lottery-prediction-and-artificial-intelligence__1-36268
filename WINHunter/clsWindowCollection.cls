VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWindowCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Window"
Attribute VB_Ext_KEY = "Member0" ,"Window"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'local variable to hold collection
Private mCol As Collection

Public Function Add(objNewMember As Form) As Integer
    'create a new object
    'Dim objNewMember As Window
    'Set objNewMember = New Window


    'set the properties passed into the method
    'objNewMember.Key = Key


On Error GoTo AddCollectionError

'    If Len(objNewMember.Tag) > 0 Then
'        mCol.Add objNewMember, UCase$(objNewMember.Name & objNewMember.Tag)
'    Else
        mCol.Add objNewMember, UCase$(objNewMember.Name)
'    End If


    'return the object created
    Add = True
    'Set objNewMember = Nothing
    Exit Function

AddCollectionError:
    Add = False
    Exit Function


End Function

Public Property Get Item(vntIndexKey As Variant) As Form
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(UCase$(vntIndexKey))
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Sub Remove(frmIn As Form)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    mCol.Remove UCase$(frmIn.Name)
    
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Public Property Get Exist(vntIndexKey As Variant) As Boolean
Dim mColItem As Form
'Set mColItem = New Window

    On Error GoTo FormNoExist
    If mCol.Count > 0 Then
        Set mColItem = mCol(UCase$(vntIndexKey))
        If mColItem.Caption <> "" Then
            Exist = True
        Else
            Exist = False
        End If
    Else
        Exist = False
    End If
    Exit Property

FormNoExist:
    Select Case Err
        Case 5
            Exist = False
        Case Else
            'Unknown error
    End Select
    Exit Property
End Property

Public Property Get Available(vntIndexKey As Variant) As Integer
Dim iNext As Integer
Dim mColItem As Form
Dim bFormFound As Boolean
Dim iCount As Integer

    For iNext = mCol.Count To 1 Step -1
        Set mColItem = mCol(iNext)
        If InStr(vntIndexKey, mColItem.Name) > 0 Then
            bFormFound = True
            iCount = iCount + 1
        Else
            'No Form Found
        End If
    Next iNext
    If Not bFormFound Then
        Available = 1
    End If
End Property

