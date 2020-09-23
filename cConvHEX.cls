VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConvHEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ======================================================================
'   Name:           cConvHEX.CLS
'   Descr:
'       Ultra-fast bin <--> hex conversion
'
'   Revision:       1.0
'   Last reviewed:  mar 24, 2004
'
'   By:             Mario Vernari
' ======================================================================
Option Explicit

    ' Assembler function stuffs
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Private asmB2H() As Byte
Private asmH2B() As Byte

Private Function AsmToByteArray(ByVal sASM As String, ByRef B() As Byte)
    ' Create a byte-array to hold the precompiled assembler code
    ReDim B(0 To Len(sASM) \ 2 - 1) As Byte

Dim I As Long
    For I = 1 To Len(sASM) Step 2
        B(I \ 2) = Val("&H" & Mid(sASM, I, 2))
    Next
End Function

Public Function BArrFromHex(sText As String, ByRef Result() As Byte) As Long
    ' Convert an HEX-string to a byte-array
Dim lNum As Long, lLength As Long
    lLength = Len(sText)
    If lLength = 0 Then Exit Function
    
    ReDim Result(Len(sText) / 2) As Byte
    CallWindowProc VarPtr(asmH2B(0)), VarPtr(Result(0)), StrPtr(sText), VarPtr(lNum), VarPtr(lLength)
    ReDim Preserve Result(lNum - 1) As Byte
    BArrFromHex = lNum
End Function

Public Function BArrToHex(BArr() As Byte, ByRef sResult As String) As String
    ' Convert a byte-array to an HEX-string
    sResult = Space(UBound(BArr) * 3 + 2)
    CallWindowProc VarPtr(asmB2H(0)), StrPtr(sResult), VarPtr(BArr(0)), UBound(BArr) + 1&, 0
End Function

Private Sub Class_initialize()
    ' Init asm funcs
    AsmToByteArray "558BEC57565053518B750C8B7D088B4D10E3278A06C0E8040C303C397E020407880783C7028A06240F0C303C397E020407880783C7044649EBD7595B585E5F8BE55DC21000", _
        asmB2H()
    
    AsmToByteArray "558BEC57565053518B750C8B7D088B5D148B1383EE024F33DBE8260000000AE4746B80FC0174F243478AC18807E812000000" & _
        "0AE4745780FC0174DEC0E0040AC18807EBD532E40BD274424A83C602668B0E67E338FEC480FD00753180F930722C" & _
        "80F939770780E10FFEC4EB2080F941721B80F946770780E937FEC4EB0F80F961720A80F966770580E957FEC4C38B4D10" & _
        "8919595B585E5F8BE55DC21000", _
        asmH2B()

End Sub
