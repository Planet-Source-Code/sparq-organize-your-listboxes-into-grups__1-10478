VERSION 5.00
Begin VB.UserControl SprqGroupList 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1530
   ScaleHeight     =   2670
   ScaleWidth      =   1530
   Begin VB.Frame Panel 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   1335
      Begin VB.VScrollBar VScroll1 
         Height          =   2055
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image IMG 
         Height          =   210
         Index           =   0
         Left            =   3180
         Top             =   2340
         Width           =   255
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   3480
         TabIndex        =   2
         Top             =   2340
         Visible         =   0   'False
         Width           =   45
      End
   End
   Begin VB.TextBox Border 
      Enabled         =   0   'False
      Height          =   495
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "SprqGroupList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Groups(999) As String
Dim Items(999) As String
Dim LabelQty As Integer
Dim GroupQty As Integer
Dim ItemQty As Integer
Dim Expand As String

Public Function AddGroup(GroupName As String)
    Groups(GroupQty) = GroupName
    GroupQty = GroupQty + 1
    RefreshList 0
End Function

Public Function AddItem(GroupName As String, ItemText As String)
  Dim Hit As Boolean
    x = 0
    Do
        If LCase(Groups(x)) = LCase(GroupName) Then Hit = True
        x = x + 1
        If Groups(x) = "" Then Exit Do
    Loop
    
    If Hit = False Then MsgBox "Group """ & GroupName & """ doesn't exsist.", vbExclamation, "Group List": Exit Function
    
    Items(ItemQty) = GroupName & "`" & ItemText
    ItemQty = ItemQty + 1
    RefreshList 0
End Function

Public Function RemoveItem(GroupName As String, ItemText As String)
  On Error GoTo Err
  Dim DelSpot As Integer
    x = 0
    Do
        If LCase(Items(x)) = LCase(GroupName) & "`" & LCase(ItemText) Then
            Items(x) = ""
            Exit Do
        Else
            x = x + 1
        End If
    Loop
    DelSpot = x
    
    For x = DelSpot To 999
        Items(x) = Items(x + 1)
        If Items(x + 1) = "" Then Exit For
    Next x
    RefreshList 0
    Exit Function

Err:
    MsgBox "Can not delete """ & ItemText & """ from Group """ & GroupName & """."
End Function

Public Sub RefreshList(Start As Integer)
  Dim MaxAmt As Integer
  On Error Resume Next
    MaxAmt = Int(Panel.Height \ 280)
    x = 0
    gc = Start
    Do
        If Groups(gc) = "" Then Exit Do
        Load Label1(x)
        Load IMG(x)
        
        Label1(x).Top = 295 * x
        IMG(x).Top = Label1(x).Top

        If Label1(x).Top >= Panel.Height Then Exit Do
        Label1(x).FontBold = True
        Label1(x).Left = 400
        IMG(x).Left = 60
        If CountItemInGroup(Groups(gc)) > 0 Then
            If Label1(x).Caption = " " & Expand Then GoTo GoodLord
            IMG(x).Picture = LoadPicture(App.Path & "\groupplus.gif")
            IMG(x).Tag = "P"
        Else
            IMG(x).Picture = LoadPicture(App.Path & "\group.gif")
            IMG(x).Tag = ""
        End If
GoodLord:
        Label1(x).Caption = " " & Groups(gc)
        Label1(x).Visible = True
        IMG(x).Visible = True
        Label1(x).Width = Panel.Width - 630
        Label1(x).Height = 210
        If Expand = "" Then GoTo 10
        If Right(Label1(x).Caption, Len(Label1(x).Caption) - 1) = Expand Then
            Dim J As Integer
            J = 0
            Do
               If Left$(Items(J), Len(Groups(gc))) = Groups(gc) Then
                    
                    x = x + 1
                    Load Label1(x)
                    Label1(x).FontBold = False
                    Label1(x).Caption = "   " & Right$(Items(J), Len(Items(J)) - (Len(Groups(gc)) + 1))
                    Label1(x).BackColor = vbWhite
                    Label1(x).ForeColor = vbBlack
                    Label1(x).Width = Panel.Width - 630
                    Label1(x).Top = (295 * x)
                    Label1(x).Left = 400
                    Label1(x).Visible = True
                    IMG(x).Visible = False
                    If x > MaxAmt Then
                        ShowScrollBar
                        VScroll1.Min = 0
                        VScroll1.Max = (x - MaxAmt) + CountGroups
                    End If
               End If
               J = J + 1
               If Items(J + 1) = "" Then GoTo 10
            Loop
        Else

        End If
10
    x = x + 1
    gc = gc + 1
    Loop
    For w = x To MaxAmt
        IMG(w).Visible = False
        Label1(w).Visible = False
    Next w
    If CountGroups > MaxAmt Then ShowScrollBar
    VScroll1.Min = 0
    VScroll1.Max = CountGroups - MaxAmt
End Sub

Function CountItemInGroup(Group As String) As Integer
  Dim ItemCount As Integer
    For x = 0 To 999
        If Items(x) = "" Then Exit For
        If LCase(Left(Items(x), Len(Group) + 1)) = LCase(Group) & "`" Then
            ItemCount = ItemCount + 1
        End If
    Next x
    CountItemInGroup = ItemCount
End Function

Function CountGroups() As Integer
    For x = 0 To 999
        If Groups(x) = "" Then CountGroups = x - 1: Exit Function
    Next x
End Function

Sub ShowScrollBar()
    VScroll1.Top = 0
    VScroll1.Left = Panel.Width - VScroll1.Width
    VScroll1.Height = Panel.Height
    VScroll1.Visible = True
    
End Sub
Private Sub IMG_DblClick(Index As Integer)
    On Error Resume Next
    Dim Um As Boolean
    For x = 0 To Label1.Count - 1
        If x = Index Then
            If IMG(x).Tag = "P" Then
                IMG(x).Picture = LoadPicture(App.Path & "\groupminus.gif")
                IMG(x).Tag = "M"
                Expand = Right(Label1(x).Caption, Len(Label1(x).Caption) - 1)
                Um = True
            ElseIf IMG(x).Tag = "M" Then
                IMG(x).Picture = LoadPicture(App.Path & "\groupplus.gif")
                IMG(x).Tag = "P"
            End If
        Else
            IMG(x).Tag = ""
        End If
    Next x
    If Um = False Then Expand = ""
    RefreshList 0
End Sub


Private Sub Label1_Click(Index As Integer)
    For x = 0 To Label1.Count - 1
        If x = Index Then
            Label1(x).BackColor = &H800000
            Label1(x).ForeColor = vbWhite
            
        Else
            Label1(x).BackColor = vbWhite
            Label1(x).ForeColor = vbbalck
        End If
    Next x
End Sub

Private Sub Label1_DblClick(Index As Integer)
    IMG_DblClick (Index)
End Sub

Private Sub UserControl_Resize()

    With Border
        .Top = 0
        .Left = 0
        .Width = Width
        .Height = Height
    End With
    With Panel
        .Top = 45
        .Left = 45
        .Width = Width - 90
        .Height = Height - 90
    End With
End Sub

Private Sub VScroll1_Change()
    RefreshList VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
    VScroll1_Change
End Sub
