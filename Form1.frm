VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "XMs / MODs / Tracks / Tunes  In VB6.0 With BassMod.Dll - Crouz Crack Me 2 By bLaCk-bytE!"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSystemCode 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtRegisterationKey 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Gen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "System Code :"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Registeration Key"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   3315
      Left            =   0
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   8220
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPathA Lib "kernel32" (ByVal nSize As Long, ByVal lpBuffer As String) As Long

Private Declare Function BASSMOD_Init Lib "bassmod.dll" (ByVal device As Long, ByVal freq As Long, ByVal flags As Long) As Integer
Private Declare Function BASSMOD_MusicLoad Lib "bassmod.dll" (ByVal mem As Integer, ByVal pfile As Any, ByVal offset As Long, ByVal Length As Long, _
        ByVal flags As Long) As Integer
Private Declare Function BASSMOD_MusicPlay Lib "bassmod.dll" () As Integer
Private Declare Function BASSMOD_MusicStop Lib "bassmod.dll" () As Integer
Private Declare Sub BASSMOD_Free Lib "bassmod.dll" ()

Dim Temp As String
Dim X As Long
Dim TempPfad As String
Dim SystemPfad As String

Private Sub cmdGenerate_Click()
    txtRegisterationKey.Text = Generate(txtSystemCode.Text)
End Sub


Private Sub Form_Load()
    'Liest das System32-Verzeichnis aus
    Temp = Space$(255)
    X = GetSystemDirectoryA(Temp, Len(Temp))
    SystemPfad = Left(Temp, X)

    'Liest das Temp-Verzeichnis aus
    Temp = Space$(255)
    X = GetTempPathA(Len(Temp), Temp)
    TempPfad = Left(Temp, X)
    
    'Beide Dateien aus den Ressourcen erstellen
    CreateFileFromRessource 101, SystemPfad & "\bassmod.dll"
    CreateFileFromRessource 102, TempPfad & "moontrip.mod"
    
     'Bassmod initialisieren und MOD laden + abspielen
    BASSMOD_Init -1, 44100, 0
    BASSMOD_MusicLoad 0, TempPfad & "moontrip.mod", 0, 0, 2 Or 512
    BASSMOD_MusicPlay
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    BASSMOD_MusicStop 'Musik stoppen
    BASSMOD_Free 'Bassmod entladen

    If Dir(TempPfad & "moontrip.mod") <> vbNullString Then 'Wenn MOD existiert
        Kill TempPfad & "moontrip.mod" 'Dann löschen
    End If
End Sub
Public Function CreateFileFromRessource(ID As Integer, FileName As String)
    Dim DataArray() As Byte

    DataArray = LoadResData(ID, "CUSTOM") 'Ressource in Array laden

    Open FileName For Binary As #1 'Datei im Binärformat öffnen/erstellen
        Put #1, , DataArray 'Und Bytes schreiben
    Close #1

    Erase DataArray
End Function

