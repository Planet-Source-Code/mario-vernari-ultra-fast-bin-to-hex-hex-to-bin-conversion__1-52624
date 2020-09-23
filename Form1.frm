VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Tahoma"
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
   ScaleHeight     =   7050
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   150
      TabIndex        =   4
      Top             =   4275
      Width           =   3165
   End
   Begin RichTextLib.RichTextBox RTF1 
      Height          =   3090
      Left            =   150
      TabIndex        =   2
      Top             =   825
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   5450
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
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
   Begin VB.CommandButton cBinToHEX 
      Caption         =   "BIN --> HEX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5775
      TabIndex        =   1
      Top             =   225
      Width           =   1515
   End
   Begin VB.TextBox tFile 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Text            =   "tFile"
      Top             =   300
      Width           =   5490
   End
   Begin VB.Label Label2 
      Caption         =   "Result:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   150
      TabIndex        =   5
      Top             =   4050
      Width           =   3765
   End
   Begin VB.Label Label1 
      Caption         =   "File to convert:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   150
      TabIndex        =   3
      Top             =   75
      Width           =   3765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Lot of thanx to Steve McMahon (vbaccelerator.com) for the cHiResTimer class.

' ------------------------------------------
' Tested on:
'   Athlon XP 2000+
'   256MB
'   Win2K server
'
'
' Average elapsed time (file size: about 1.4 MB)
'   bin --> hex     45 ms
'   hex --> bin     84 ms
'
'
' Average rates:
'   bin --> hex     28.8 MBytes/s
'   hex --> bin     15.7 MBytes/s
' ------------------------------------------

Private cCvHEX As cConvHEX

Private Sub cBinToHEX_Click()
    ' Read file
    List1.Clear
    
    On Error GoTo Label_Error
    
Dim B_In() As Byte, B_Out() As Byte
    Open tFile.Text For Binary As #1
        ReDim B_In(LOF(1)) As Byte
        Get #1, , B_In
    Close #1
    
    List1.AddItem "File size:" & UBound(B_In) + 1
    
    ' Start conversion: BIN --> HEX
Dim cTimer As cHiResTimer
    Set cTimer = New cHiResTimer
    cTimer.StartTimer
Dim sHEX As String
    cCvHEX.BArrToHex B_In, sHEX
'    sHEX = cCvHEX.BArrToHex(B_In)
    cTimer.StopTimer
    
    List1.AddItem "HEX size: " & Len(sHEX)
    List1.AddItem "Conv time: " & Int(cTimer.ElapsedTime * 1000) & " ms"
    
    If cTimer.ElapsedTime Then
        List1.AddItem "Rate: " & Format(UBound(B_In) / cTimer.ElapsedTime / 1048576#, "#0.0") & " MBytes/s"
    Else
        List1.AddItem "Rate: n/a"
    End If
    List1.AddItem "----------------"
    
    ' Display the resulting stream
    '   NOTE: due to large memory allocation,
    '   it's better to limit the displayed stream
    RTF1.Text = Left(sHEX, 10000)
    
    ' Start conversion: HEX --> BIN
    Set cTimer = New cHiResTimer
    cTimer.StartTimer
    cCvHEX.BArrFromHex sHEX, B_Out
    cTimer.StopTimer
    
    List1.AddItem "Bin size: " & UBound(B_Out) + 1
    List1.AddItem "Conv time: " & Int(cTimer.ElapsedTime * 1000) & "ms"
    
    If cTimer.ElapsedTime Then
        List1.AddItem "Rate: " & Format(UBound(B_Out) / cTimer.ElapsedTime / 1048576#, "#0.0") & " MBytes/s"
    Else
        List1.AddItem "Rate: n/a"
    End If
    List1.AddItem "----------------"
    
    ' Let's verify the resulting stream with the original one
Dim sTestIN As String, sTestOUT As String
    sTestIN = B_In
    sTestOUT = B_Out
    List1.AddItem "Verify: " & IIf(sTestIN = sTestOUT, "Ok", "Fail")
    Exit Sub
    
    
Label_Error:
    MsgBox Err.Description, vbExclamation, App.Title
End Sub

Private Sub Form_Load()
    Set cCvHEX = New cConvHEX
    
    tFile.Text = "\winnt\system32\msvbvm60.dll"
    RTF1.Text = ""
End Sub
