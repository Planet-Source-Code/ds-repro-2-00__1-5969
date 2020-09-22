VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{C7978AF1-C807-11D3-8138-C648DCBB7330}#9.0#0"; "DsReproductor.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPlayer 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ds Mp3"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4140
   Icon            =   "Repro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Repro.frx":0442
   ScaleHeight     =   4065
   ScaleWidth      =   4140
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   1635
      Top             =   2325
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Abrir Archivo"
      Height          =   315
      Left            =   1425
      TabIndex        =   8
      Top             =   3630
      Width           =   1260
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Limpiar Lista"
      Height          =   315
      Left            =   75
      TabIndex        =   7
      Top             =   3630
      Width           =   1260
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Borrar Tema"
      Height          =   315
      Left            =   1425
      TabIndex        =   5
      Top             =   3255
      Width           =   1260
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar lista"
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   3255
      Width           =   1260
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   345
      Left            =   3120
      TabIndex        =   3
      Top             =   3255
      Width           =   930
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   975
      Top             =   2565
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   525
      Top             =   2610
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   45
      TabIndex        =   2
      Top             =   1755
      Width           =   4005
   End
   Begin PicClip.PictureClip PicClip1 
      Left            =   660
      Top             =   1875
      _ExtentX        =   3598
      _ExtentY        =   953
      _Version        =   393216
      Rows            =   2
      Cols            =   6
      Picture         =   "Repro.frx":17BB6
   End
   Begin DsReproductor.DsRepro DsRepro1 
      Left            =   600
      Top             =   2340
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   165
      TabIndex        =   6
      Top             =   975
      Width           =   3780
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "1999.-Carlos D'Agostino. - Todos los Derechos. "
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label LCD 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   555
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1290
   End
   Begin VB.Image Image1 
      Height          =   285
      Index           =   5
      Left            =   2070
      Top             =   1335
      Width           =   330
   End
   Begin VB.Image Image1 
      Height          =   285
      Index           =   4
      Left            =   1605
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   285
      Index           =   3
      Left            =   1260
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   285
      Index           =   2
      Left            =   915
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   270
      Index           =   1
      Left            =   570
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   0
      Left            =   240
      Top             =   1320
      Width           =   360
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Playing As Boolean, PlayIndex As Integer, stoping As Boolean
Dim a As String, bTime As Boolean

Const conHwndTopmost = -1
Const conHwndNoTopmost = -2
Const conSwpNoActivate = &H10
Const conSwpShowWindow = &H40


Private Sub Command1_Click()
    frmPlayer.Height = 2115
    If List1.ListCount > 0 Then
       SetControls True, True, True, True, True, True
    Else
       SetControls False, False, False, False, False, True
    End If
End Sub

Private Sub Command2_Click()
  Dim archi As String, i As Integer
  Static LastFilter
  If LastFilter = 0 Then LastFilter = 1
  On Error GoTo ErrHandler
  Dialog1.CancelError = True
  Dialog1.InitDir = App.Path
  Dialog1.DialogTitle = "Grabar Listas de Sonidos"
  Dialog1.Flags = cdlOFNHideReadOnly
  Dialog1.Filter = "Playlists *.m3u|*.M3U;"
  Dialog1.FilterIndex = LastFilter
  Dialog1.FileName = ""
  Dialog1.ShowSave
        
  archi = Dialog1.FileName
  Open archi For Output As #1
  List1.ListIndex = 0
  For i = 1 To List1.ListCount
      Print #1, List1.Text
      List1.ListIndex = List1.ListIndex + 1
  Next i
  Close #1
ErrHandler:
    Exit Sub

End Sub

Private Sub Command3_Click()
  If List1.ListCount > 0 Then
     List1.RemoveItem List1.ListIndex
  End If
End Sub

Private Sub Command4_Click()
  List1.Clear
End Sub

Private Sub Command5_Click()
  a = DsRepro1.AbrirSonido
  Dim b As String
  If Right(a, 3) <> "M3U" And Right(a, 3) <> "m3u" Then
     List1.AddItem a
  Else
     Close #1
     Open a For Input As #1
     Do While Not EOF(1)
        Input #1, b
        List1.AddItem b
     Loop
     Close #1
 End If
 SetControls True, True, True, True, True, False
End Sub

Private Sub DsRepro1_Reproduciendo(tiempo As String)
  LCD.Caption = tiempo
End Sub

Private Sub DsRepro1_Termine()
  On Error GoTo errores
  If Playing = True Then
     List1.ListIndex = List1.ListIndex + 1
     If List1.ListIndex < List1.ListCount Then
        DsRepro1.Reproducir (List1.Text)
        Label2.Caption = List1.Text
     End If
  End If
    Exit Sub
errores:
End Sub

Private Sub Form_Load()
  SetControls False, False, False, False, False, True
  frmPlayer.Height = 2115
  SetWindowPos hWnd, conHwndTopmost, 0, 0, 280, 135, conSwpNoActivate Or conSwpShowWindow
  
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  Select Case Index
         Case 0
              bTime = False
              PicClip1.ClipX = 0
              PicClip1.ClipY = 18
              PicClip1.ClipWidth = 23
              PicClip1.ClipHeight = 15
              Image1(Index).Picture = PicClip1.Clip
              Timer2.Enabled = True
         Case 1
              PicClip1.ClipX = 23
              PicClip1.ClipY = 18
              PicClip1.ClipWidth = 24
              PicClip1.ClipHeight = 15
              Image1(Index).Picture = PicClip1.Clip
              Playing = True
              PlayIndex = 0
              stoping = False
              If List1.Text = "" Then
                 If List1.ListIndex <> -1 Then
                    List1.ListIndex = 0
                 End If
              End If
              DsRepro1.Reproducir (List1.Text)
              Label2.Caption = List1.Text
         Case 2
              PicClip1.ClipX = 46
              PicClip1.ClipY = 18
              PicClip1.ClipWidth = 24
              PicClip1.ClipHeight = 15
              Image1(Index).Picture = PicClip1.Clip
         Case 3
              PicClip1.ClipX = 69
              PicClip1.ClipY = 18
              PicClip1.ClipWidth = 24
              PicClip1.ClipHeight = 15
              Image1(Index).Picture = PicClip1.Clip
              stoping = True
         Case 4
              bTime = False
              PicClip1.ClipX = 92
              PicClip1.ClipY = 18
              PicClip1.ClipWidth = 23
              PicClip1.ClipHeight = 15
              Image1(Index).Picture = PicClip1.Clip
              Timer1.Enabled = True
         Case 5
              PicClip1.ClipX = 114
              PicClip1.ClipY = 16
              PicClip1.ClipWidth = 20
              PicClip1.ClipHeight = 15
              Image1(Index).Picture = PicClip1.Clip
  End Select
End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim i As Integer
  Select Case Index
         Case 0
              PicClip1.ClipX = 0
              PicClip1.ClipY = 0
              PicClip1.ClipWidth = 24
              PicClip1.ClipHeight = 15
              Image1(Index).Picture = PicClip1.Clip
              stoping = True
              PlayIndex = PlayIndex - 2
              If PlayIndex < 0 Then PlayIndex = 0
              PlayIndex = PlayIndex + 1
              Timer2.Enabled = False
              If bTime = False Then
                 If List1.ListIndex <> 0 Then
                    If List1.ListIndex <> -1 Then
                       List1.ListIndex = List1.ListIndex - 2
                    End If
                 Else
                    List1.ListIndex = 0
                 End If
                 Label2.Caption = List1.Text
                 DsRepro1_Termine
              End If
         Case 1
              PicClip1.ClipX = 22
              PicClip1.ClipY = 0
              PicClip1.ClipWidth = 22
              PicClip1.ClipHeight = 10
              Image1(1).Picture = PicClip1.Clip
              PlayIndex = PlayIndex + 1
              SetControls True, True, True, True, True, True
         Case 2
              PicClip1.ClipX = 45
              PicClip1.ClipY = 0
              PicClip1.ClipWidth = 24
              PicClip1.ClipHeight = 15
              Image1(Index).Picture = PicClip1.Clip
              DsRepro1.Pausa
         Case 3
              PicClip1.ClipX = 68
              PicClip1.ClipY = 0
              PicClip1.ClipWidth = 23
              PicClip1.ClipHeight = 15
              Image1(Index).Picture = PicClip1.Clip
              PicClip1.ClipX = 22
              PicClip1.ClipY = 0
              PicClip1.ClipWidth = 22
              PicClip1.ClipHeight = 10
              Image1(1).Picture = PicClip1.Clip
              Playing = False
              DsRepro1.Parar
              SetControls True, True, False, False, True, True
         Case 4
              PicClip1.ClipX = 91
              PicClip1.ClipY = 0
              PicClip1.ClipWidth = 23
              PicClip1.ClipHeight = 15
              Image1(Index).Picture = PicClip1.Clip
              stoping = True
              PlayIndex = PlayIndex + 1
              Timer1.Enabled = False
              If bTime = False Then
                 DsRepro1_Termine
              End If
         Case 5
              PicClip1.ClipX = 115
              PicClip1.ClipY = 0.5
              PicClip1.ClipWidth = 20
              PicClip1.ClipHeight = 15
              Image1(Index).Picture = PicClip1.Clip
              frmPlayer.Height = 4440
     End Select
End Sub

Private Sub SetControls(ByVal bRev As Boolean, ByVal bPlay As Boolean, _
    ByVal bPause As Boolean, ByVal bStop As Boolean, ByVal bForw As Boolean, ByVal bEject As Boolean)
  Image1(0).Enabled = bRev
  Image1(1).Enabled = bPlay
  Image1(2).Enabled = bPause
  Image1(3).Enabled = bStop
  Image1(4).Enabled = bForw
  Image1(5).Enabled = bEject
End Sub

Private Sub Timer1_Timer()
    DsRepro1.Adelantar
    bTime = True
End Sub

Private Sub Timer2_Timer()
    DsRepro1.Atrasar
    bTime = True
End Sub
