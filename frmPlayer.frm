VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPlayer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MP3Player"
   ClientHeight    =   1335
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3660
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3660
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox ProgressBar 
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   150
      ScaleHeight     =   180
      ScaleWidth      =   3255
      TabIndex        =   6
      Top             =   900
      Width           =   3315
   End
   Begin VB.Frame ToolBar 
      Caption         =   "Controles"
      Height          =   690
      Left            =   150
      TabIndex        =   0
      Top             =   75
      Width           =   1815
      Begin VB.CommandButton Botones 
         DisabledPicture =   "frmPlayer.frx":0442
         Height          =   315
         Index           =   4
         Left            =   1350
         Picture         =   "frmPlayer.frx":0514
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   255
         Width           =   315
      End
      Begin VB.CommandButton Botones 
         DisabledPicture =   "frmPlayer.frx":05DE
         Height          =   315
         Index           =   3
         Left            =   1050
         Picture         =   "frmPlayer.frx":06A8
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   255
         Width           =   315
      End
      Begin VB.CommandButton Botones 
         DisabledPicture =   "frmPlayer.frx":074A
         Height          =   315
         Index           =   2
         Left            =   750
         Picture         =   "frmPlayer.frx":0824
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   255
         Width           =   315
      End
      Begin VB.CommandButton Botones 
         DisabledPicture =   "frmPlayer.frx":08EE
         Height          =   315
         Index           =   1
         Left            =   450
         Picture         =   "frmPlayer.frx":09C8
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   255
         Width           =   315
      End
      Begin VB.CommandButton Botones 
         DisabledPicture =   "frmPlayer.frx":0A92
         Height          =   315
         Index           =   0
         Left            =   150
         Picture         =   "frmPlayer.frx":0B6C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   270
         Width           =   315
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Reproducir archivo..."
      Filter          =   "MPEG Layer III (*.mp3)|*.mp3|"
      Flags           =   4
      InitDir         =   "d:\mp3"
   End
   Begin VB.Label Texto 
      Caption         =   "Germán Oltra             <goltra@geocities.com>"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   1
      Left            =   150
      TabIndex        =   9
      Top             =   2850
      Width           =   3315
   End
   Begin VB.Label Texto 
      Caption         =   $"frmPlayer.frx":0C46
      Height          =   1290
      Index           =   0
      Left            =   150
      TabIndex        =   8
      Top             =   1500
      Width           =   3315
   End
   Begin VB.Line Separador 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   3675
      X2              =   0
      Y1              =   1365
      Y2              =   1365
   End
   Begin VB.Line Separador 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   3675
      Y1              =   1350
      Y2              =   1350
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
      Height          =   540
      Left            =   2100
      TabIndex        =   7
      Top             =   225
      Width           =   1365
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'====================================================================
'                 --------------------------------
'                       M P 3 - P L A Y E R
'                 --------------------------------
'
'   - Reproductor *experimental* de archivos MPEG Layer III, usando
'     la librería mp3.dll y empleando técnicas de  MultiThreading y
'     Subclasificación... (vaya tela! :-)
'
'
' Germán Oltra                                           ICQ#11765971
' http:\\goltra.home.ml.org                    <goltra@geocities.com>
'====================================================================

Private Total As Double
Private PerFrame As Double
Private ProgressTop As Integer
Private ProgressValue As Integer

Private Playing As Boolean
'Gestiona las acciones sobre los controles
Private Sub botones_Click(Index As Integer)
    Select Case Index
        Case 0
            PlayFile
            SetControls False, True, True, True, False
            Playing = True
        Case 1, 2: MsgBox "Las funciones de Seek no las he implementado, eso os lo dejo a vosotros! :-)"
        Case 3
            StopFile
            SetControls True, False, False, False, True
            ProgressValue = 0
            ProgressBar.AutoRedraw = True
            ProgressBar.Cls
            ProgressBar.AutoRedraw = False
            LCD.Caption = InTime(Total)
            Playing = False
        Case 4
            Dialog.ShowOpen
            If Dialog.FileName = "" Then Exit Sub
            Mp3Info.FileName = Dialog.FileName
            Me.Caption = GetFileName(Dialog.FileName)
            SetControls True, False, False, False, True
            GetTime Total, PerFrame
            ProgressTop = Fix((Total / PerFrame) / 16)
            ProgressValue = 0
            LCD.Caption = InTime(Total)
    End Select
End Sub
'Secuencia de inicialización
Private Sub Form_Load()
    HookForm Me
    Mp3Info.Owner = Me.hWnd
    Playing = False
    SetControls False, False, False, False, True
End Sub
'Detener la interceptación de mensajes
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    unHookForm
End Sub
'No puedo salir si se esta reproduciendo
Private Sub Form_Unload(Cancel As Integer)
    If Playing Then
        StopFile
        Playing = False
        Cancel = 1
    End If
End Sub
'Activa-Desactiva los botones
Private Sub SetControls(ByVal bPlay As Boolean, ByVal bRewind As Boolean, _
    ByVal bForward As Boolean, ByVal bStop As Boolean, ByVal bEject As Boolean)
    Botones(0).Enabled = bPlay
    Botones(1).Enabled = bRewind
    Botones(2).Enabled = bForward
    Botones(3).Enabled = bStop
    Botones(4).Enabled = bEject
End Sub
'Avanza la barra de progreso
Private Sub PushBar()
    Dim Contador As Integer
    ProgressBar.AutoRedraw = True
    For Contador = 1 To Fix(ProgressBar.ScaleWidth / ProgressTop)
        ProgressBar.Line (ProgressValue + Contador, 0)-Step(0, ProgressBar.ScaleHeight)
    Next Contador
    ProgressBar.AutoRedraw = False
    ProgressValue = ProgressValue + Contador
End Sub
'Destino de los mensajes interceptados
Public Sub MessageReceived(ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long)
    Select Case Message
        Case FRAME_POS: LCD.Caption = InTime(Total - (PerFrame * wParam)): PushBar
        Case APPLY_POS: Debug.Print Now, "Seek Process Done!"
        Case PLAY_STOP: SetControls True, False, False, False, True: Playing = False
    End Select
End Sub
'Muestra o esconde el *about*
Private Sub LCD_Click()
    frmPlayer.Height = IIf(frmPlayer.Height = 3615, 1710, 3615)
End Sub
