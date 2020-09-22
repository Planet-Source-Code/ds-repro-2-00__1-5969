VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl DsRepro 
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1650
   InvisibleAtRuntime=   -1  'True
   Picture         =   "DsRepro.ctx":0000
   ScaleHeight     =   990
   ScaleWidth      =   1650
   ToolboxBitmap   =   "DsRepro.ctx":066A
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   225
      Top             =   465
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   512
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   1095
      Top             =   480
   End
End
Attribute VB_Name = "DsRepro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private MediaControl As IMediaControl
Private MediaPosition As IMediaPosition
Private BasicAudio As IBasicAudio
Private SongLen As Double, Elapsed As Double
Private Paused As Boolean
Private OptDefPath As String
Private Prg As String, Sect As String
Private AddName As String, AddPath As String


Event Termine()
Event Reproduciendo(tiempo As String)
Public Autor As String

Public Function Reproducir(Archivo As String)
  On Error GoTo errores
  Set MediaControl = New FilgraphManager
  Set MediaPosition = MediaControl
  Set BasicAudio = MediaControl

  If Archivo <> "" Then
     MediaControl.RenderFile (Archivo)
     MediaPosition.CurrentPosition = 0
     SongLen = MediaPosition.Duration
     MediaControl.Run
     Timer1.Enabled = True
  End If
Exit Function
errores:
  MsgBox "El Control permite un solo Objeto en cada Formulario o Archivo Inválido", vbCritical, "DsRepro"
End Function

Public Function Parar()
  MediaControl.Stop
  Timer1.Enabled = False
  Set MediaControl = Nothing
  Set MediaPosition = Nothing
  Set BasicAudio = Nothing
  RaiseEvent Termine
End Function

Public Function Adelantar()
  Dim p As Double
  p = Elapsed + 10
  Call Recu(p)
End Function

Public Function Atrasar()
  Dim p As Double
  p = Elapsed - 10
  Call Recu(p)
End Function

Public Function Pausa()
  If Paused = True Then
     MediaControl.Run
  Else
     MediaControl.Pause
  End If
  Paused = Not Paused
End Function
Private Sub Timer1_Timer()
  Dim pos As Single, min As Integer, sec As Integer
  Dim cPos As Double, tiempo As String
  

  Elapsed = MediaPosition.CurrentPosition
  
  cPos = Elapsed
  min = Int(cPos / 60): sec = Int(cPos - min * 60)
  tiempo = Format(min, "00") + ":" + Format(sec, "00")
  RaiseEvent Reproduciendo(tiempo)
  
  If Elapsed >= SongLen Then GoSub GoNext
  Exit Sub
GoNext:
  MediaControl.Stop
  Timer1.Enabled = False
  Set MediaControl = Nothing
  Set MediaPosition = Nothing
  Set BasicAudio = Nothing
  RaiseEvent Termine

End Sub

Private Sub UserControl_Initialize()
 Set MediaControl = New FilgraphManager
 Set MediaPosition = MediaControl
 Set BasicAudio = MediaControl
 Paused = False
 Autor = "2000.-Carlos D´Agostino."
 Prg = "vbampprodx": Sect = "config"
 OptDefPath = GetSetting(Prg, Sect, "Path", "")
End Sub

Private Sub UserControl_Resize()
  Size 32 * Screen.TwipsPerPixelX, 32 * Screen.TwipsPerPixelY
End Sub

Public Function TiempoTotal() As String
  Dim cPos As Single, min As Integer, sec As Integer
  cPos = SongLen
  min = Int(cPos / 60): sec = Int(cPos - min * 60)
  TiempoTotal = Format(min, "00") + ":" + Format(sec, "00")
End Function

Private Sub Recu(ByVal SPos As Long)
If MediaControl Is Nothing Then Exit Sub
    If SPos < 0 Then SPos = 0
    If SPos > SongLen Then SPos = SongLen - 1
    MediaPosition.CurrentPosition = SPos
End Sub

Function AbrirSonido() As String
    Static LastFilter
    If LastFilter = 0 Then LastFilter = 1
    On Error GoTo ErrHandler
    CommonDialog1.CancelError = True
    CommonDialog1.InitDir = OptDefPath
    CommonDialog1.DialogTitle = "Abrir archivo de Audio"
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.Filter = "MPEG Audio Files|*.MP?|ActiveMovie Files|*.MP?;*.MPEG;*.DAT;*.WAV;*.AU;*.MID;*.RMI;*.AIF?;*.MOV;*.QT;*.AVI;*.M1V;*.RA;*.RAM;*.RM;*.RMM|Music Modules|*.MOD;*.MTM;*FAR;*.669;*.OKT;*.STM;*.S3M;*.NST;*.WOW;*.XM|Playlists|*.M3U;*.PLS"
    CommonDialog1.FilterIndex = LastFilter
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
        
    AbrirSonido = CommonDialog1.FileName
    OptDefPath = Left(AddPath, Len(AddPath) - Len(AddName))
    SaveSetting Prg, Sect, "Path", OptDefPath
    
ErrHandler:
    Exit Function
End Function

