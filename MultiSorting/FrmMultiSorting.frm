VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMultiSorting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MultiSorting In ListView"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   Icon            =   "FrmMultiSorting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   405
      Left            =   4395
      TabIndex        =   2
      Top             =   3450
      Width           =   2025
   End
   Begin MSComctlLib.ListView lvwQuery 
      Height          =   2580
      Left            =   45
      TabIndex        =   0
      Top             =   660
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   4551
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "imglstListImages"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imglstListImages 
      Left            =   135
      Top             =   3285
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMultiSorting.frx":030A
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMultiSorting.frx":0464
            Key             =   "Down"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblKey 
      Alignment       =   2  'Center
      Caption         =   "Hold Ctrl Key for Multi Selection"
      Height          =   330
      Left            =   2835
      TabIndex        =   4
      Top             =   315
      Width           =   4905
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   9885
      TabIndex        =   3
      Top             =   3540
      Width           =   705
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Multi Sorting With ListView"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   195
      TabIndex        =   1
      Top             =   15
      Width           =   10260
   End
End
Attribute VB_Name = "FrmMultiSorting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConn As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim oTempRs As New ADOR.Recordset

Private Sub Form_Unload(Cancel As Integer)
 Call cmdClose_Click
End Sub

Private Sub lblInfo_Click()
    MsgBox "Please feel free to write your Comments/Suggestions. Thnx!" & vbCrLf & "-Deepakk_2k@yahoo.com"
End Sub

Private Sub lblInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo.FontUnderline = True
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo.FontUnderline = False
End Sub
Private Sub cmdClose_Click()
    Set oRs = Nothing
    Set oConn = Nothing
    Set oTempRs = Nothing
    End
End Sub

Private Sub Form_Load()

On Error GoTo ErrHnd
    Dim strConn As String, i As Byte
        
    '*** If you dont have Jet OLEDB 4.0 driver Use this Connection ***
    'strConn = "Provider=Microsoft.Jet.OLEDB.3.51;" & _
    "Data Source=" & App.Path & "\Users.mdb;"
    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & App.Path & "\Users.mdb;"
    
    oConn.Open strConn
    oRs.CursorLocation = adUseClient
    oRs.Open "Select * from [Users]", oConn, adOpenStatic, adLockOptimistic
    
    With lvwQuery.ColumnHeaders
    .Clear
            '*** Creating Header of ListView ***
   For i = 1 To oRs.Fields.Count
     .Add i, "F" & i, oRs.Fields.Item(i - 1).Name
    '* Creating Tmp RecordSet *
    oTempRs.Fields.Append oRs.Fields.Item(i - 1).Name, oRs.Fields.Item(i - 1).Type, oRs.Fields.Item(i - 1).DefinedSize, adFldIsNullable
   Next i
End With
    oTempRs.CursorLocation = adUseClient
    oTempRs.CursorType = adOpenStatic
    oTempRs.Open
    
         With lvwQuery.ListItems
Do Until oRs.EOF
            .Add , "Z" & oRs(0), oRs(0)
            oTempRs.AddNew
            oTempRs.Fields(0).Value = oRs(0)
    For i = 2 To oRs.Fields.Count
     If IsNull(oRs(i - 1)) = False Then
            .Item("Z" & oRs(0)).ListSubItems.Add , "K" & i, oRs(i - 1)
            oTempRs.Fields(i - 1).Value = oRs(i - 1)
     Else
            .Item("Z" & oRs(0)).ListSubItems.Add , "K" & i, ""
            oTempRs.Fields(i - 1).Value = Null
     End If
    Next i
    oRs.MoveNext
Loop
            End With
             oTempRs.Update
             
    Exit Sub
ErrHnd:
    
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description
End Sub
Private Sub lvwQuery_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Static SortItem() As String, SortOrder() As Boolean
Static Clicked As Boolean, SingleChecked As Boolean
Dim i As Byte, j As Integer

On Error GoTo ErrHnd
    If KeysPressed(vbKeyControl) = True Then
     If Clicked = False Then '*** clicked first time ***
        Clicked = True
        ReDim SortItem(0): ReDim SortOrder(0)
        SortItem(0) = ColumnHeader: SortOrder(0) = True
        
     Else   '*** Multi Selection started ***
                j = InColletion(ColumnHeader, SortItem)
            If j >= 0 Then
                '* Item Exit in the List, Only to Change the Sorting Order *
                SortOrder(j) = IIf(SortOrder(j) = True, False, True)
            Else
                '* Inserting New Item to Sort *
                ReDim Preserve SortItem(UBound(SortItem) + 1)
                ReDim Preserve SortOrder(UBound(SortOrder) + 1)
                
                SortItem(UBound(SortItem)) = ColumnHeader
                SortOrder(UBound(SortOrder)) = True
            End If
        
     End If
        
    Else '* Reset the Sorting *
            '*** Asc/Desc with out Addition Key ***
      If SingleChecked = False Then '* First Time assine New values *
        ReDim SortItem(0): ReDim SortOrder(0)
        SortItem(0) = ColumnHeader: SortOrder(0) = True
        SingleChecked = True
      Else
          '*Checking whether the item clecked twice or not*
          If SortItem(0) = ColumnHeader Then
            SortOrder(0) = IIf(SortOrder(0) = True, False, True)
          Else
            ReDim SortItem(0): ReDim SortOrder(0)
            SortItem(0) = ColumnHeader: SortOrder(0) = True
          End If
      End If
      
        Clicked = False
    End If
 MultiSort_ListView SortItem, SortOrder
 'MsgBox ColumnHeader
 Exit Sub
ErrHnd:
    MsgBox "Error: " & Err.Number & Err.Description
End Sub

'*** Filling Listview - Multisorting ***
Private Sub MultiSort_ListView(Fields As Variant, Orders As Variant)
    Dim i As Integer, j As Integer
    Dim KeepHeader_Order() As String
    Dim strOrderBy As String
    
      For i = 0 To UBound(Fields)
        strOrderBy = strOrderBy & "[" & Fields(i) & "] " & _
                     IIf(Orders(i) = True, "ASC", "DESC") & ","
      Next i
      strOrderBy = Left(strOrderBy, Len(strOrderBy) - 1) '* Removing "," from the Query String *
      Debug.Print strOrderBy
    
    oTempRs.Sort = strOrderBy
    
  ' *** Reading Current Header Column Order ***
  ReDim KeepHeader_Order(lvwQuery.ColumnHeaders.Count - 1)
   For i = 0 To lvwQuery.ColumnHeaders.Count - 1
     KeepHeader_Order((lvwQuery.ColumnHeaders.Item(i + 1).Position) - 1) = lvwQuery.ColumnHeaders.Item(i + 1).Text
   Next i
  
            '*Cleaning *
    lvwQuery.ColumnHeaders.Clear
    lvwQuery.ListItems.Clear
    
  With lvwQuery.ColumnHeaders
    
   '*** filling ListView With Temp RecordSet ***
   '* Header Only *'
    For i = 0 To UBound(KeepHeader_Order)
    For j = 0 To UBound(Fields)

        If Fields(j) = KeepHeader_Order(i) Then
        '*Setting the Up and Down Icons on the ListView *
.Add i + 1, "F" & i, KeepHeader_Order(i), , , IIf(Orders(j) = True, "Up", "Down")
            Exit For
        End If
    Next j
        If j = UBound(Fields) + 1 Then .Add i + 1, "F" & i, KeepHeader_Order(i)
   Next i
  End With
  
    '* ListView Items *'
    
 oTempRs.MoveFirst
         With lvwQuery.ListItems
Do Until oTempRs.EOF

            .Add , "Z" & oTempRs(0), oTempRs(KeepHeader_Order(0))
     For i = 1 To UBound(KeepHeader_Order)
          If Not IsNull(oTempRs(KeepHeader_Order(i))) Then
            .Item("Z" & oTempRs(0)).ListSubItems.Add , "K" & i, oTempRs(KeepHeader_Order(i))
          Else
           .Item("Z" & oTempRs(0)).ListSubItems.Add , "K" & i, ""
          End If
    Next i
    oTempRs.MoveNext
Loop
        End With
  
End Sub
' *** Finding where a string belongs to an Array or not ***
Private Function InColletion(ByVal SearchStr As String, TheCollection As Variant) As Integer
    Dim i As Byte, Result As Integer
     Result = -1
     
     For i = 0 To UBound(TheCollection)
        If SearchStr = TheCollection(i) Then
            Result = i: Exit For
        End If
     Next i
     
  InColletion = Result
End Function

