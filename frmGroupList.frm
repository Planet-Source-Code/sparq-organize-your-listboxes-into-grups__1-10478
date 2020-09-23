VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2730
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   2730
   StartUpPosition =   3  'Windows Default
   Begin Project1.SprqGroupList SprqGroupList1 
      Height          =   2355
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   4154
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove ""Hi"" From Group 1"
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   3060
      Width           =   2475
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add Item to Group 1"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   2460
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Group"
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   2460
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   3780
      Width           =   45
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Static Count As Integer
    Count = Count + 1
    SprqGroupList1.AddGroup "Group " & Count
End Sub

Private Sub Command2_Click()
  SprqGroupList1.AddItem "Group 2", "Yo"
End Sub

Private Sub Command3_Click()
    SprqGroupList1.RemoveItem "Group 1", "Hi"
End Sub

Private Sub Form_Load()
    Label1 = "This code will show you how to" & vbCrLf & _
             "make a listbox with groups. Even" & vbCrLf & _
             "though this code is not working" & vbCrLf & _
             "100%, it is MORE than enough to" & vbCrLf & _
             "start you in making your own!" & vbCrLf & _
             "There are hundreds of uses for" & vbCrLf & _
             "Code like this. If you like, " & vbCrLf & _
             "Please Vote!" & vbCrLf & vbCrLf & _
             "<jay@alphamedia.net>"
End Sub
