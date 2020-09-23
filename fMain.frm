VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Read (text)file via API calls"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   317
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReadClass 
      Caption         =   "Read file via class"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   4320
      Width           =   2220
   End
   Begin RichTextLib.RichTextBox txtFileContents 
      Height          =   3855
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6800
      _Version        =   393217
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"fMain.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read file via module"
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   4320
      Width           =   2220
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status: Click the button :-)"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   4020
      Width           =   2310
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oTextRead As clsTextRead
Attribute oTextRead.VB_VarHelpID = -1

Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Private Sub cmdRead_Click()
   Dim fSucces  As Boolean
   Dim tVbPath As String
   
   ' Update button
   cmdRead.Enabled = False
   cmdReadClass.Enabled = False
   lblStatus.Caption = "Status: Wait ..."
   lblStatus.Refresh
   
   ' On my operating system, this is the path to VB 6
   tVbPath = "C:\PROGRA~1\MICROS~2\VB98\"
   txtFileContents.Text = ReadTextFile(tVbPath & "redist.txt", fSucces)
   
   ' "fSucces" holds the information if the file was read succesfully...
   If fSucces Then
      lblStatus.Caption = "Status: Succes, took " & CStr(Val(mlTimeTook)) & " ms, at " & CStr(mlSpeed) & " mb/sec"
   Else
      lblStatus.Caption = "Status: Error"
   End If

   ' Update button
   cmdRead.Enabled = True
   cmdReadClass.Enabled = True
   lblStatus.Refresh
End Sub

Private Sub cmdReadClass_Click()
   Dim tReturn As String
   Dim tVbPath As String
   Dim lTimeStart As Long
   Dim lTimeEnd   As Long
   Dim lTimeTook  As Long
   Dim lSpeed     As Long
   
   '----------------------------------------
   ' Update button
   '----------------------------------------
   cmdRead.Enabled = False
   cmdReadClass.Enabled = False
   lblStatus.Caption = "Status: Wait ..."
   lblStatus.Refresh
   
   '----------------------------------------
   ' Create class, and try to read file
   '----------------------------------------
   Set oTextRead = New clsTextRead
   tVbPath = "C:\PROGRA~1\MICROS~2\VB98\"
   oTextRead.FileName = tVbPath & "redist.txt"
   If (oTextRead.OpenFile = True) Then
      lTimeStart = GetTickCount
      Do While (oTextRead.FileAtEOF = False)
         tReturn = tReturn & oTextRead.ReadBuffer
      Loop
      oTextRead.CloseFile
   End If
   lTimeEnd = GetTickCount
   lTimeTook = lTimeEnd - lTimeStart
   On Error Resume Next
   lSpeed = (Len(tReturn) / lTimeTook) / 1024
   Err.Clear
   
   Set oTextRead = Nothing
   txtFileContents.Text = tReturn
   
   If (tReturn <> "") Then
      lblStatus.Caption = "Status: Succes, took " & CStr(Val(lTimeTook)) & " ms, at " & CStr(lSpeed) & " mb/sec"
   Else
      lblStatus.Caption = "Status: Error"
   End If
   
   '----------------------------------------
   ' Update button
   '----------------------------------------
   cmdRead.Enabled = True
   cmdReadClass.Enabled = True
   lblStatus.Refresh
End Sub
