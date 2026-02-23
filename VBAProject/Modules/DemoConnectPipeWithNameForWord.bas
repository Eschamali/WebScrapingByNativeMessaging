Attribute VB_Name = "DemoConnectPipeWithNameForWord"
'***************************************************************************************************
'        PowerAutomateDesktop と PAD.BrowserNativeMessageHost.exe のやり取りの核と言える
'　「名前付きパイプ（クライアント化）」による双方向なやり取りの概念実証デモンストレーションです
'
' このコードは、Excel以外のVBA(クライアント)用です。WordやAccessなどのVBAに貼り付けて順番に実行してください
'***************************************************************************************************
Option Explicit



' --- 名前付きパイプ用 WinAPI 宣言 ---
#If VBA7 Then
    Private Declare PtrSafe Function CreateNamedPipe Lib "kernel32" Alias "CreateNamedPipeA" (ByVal lpName As String, ByVal dwOpenMode As Long, ByVal dwPipeMode As Long, ByVal nMaxInstances As Long, ByVal nOutBufferSize As Long, ByVal nInBufferSize As Long, ByVal nDefaultTimeOut As Long, ByVal lpSecurityAttributes As LongPtr) As LongPtr
    Private Declare PtrSafe Function ConnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As LongPtr, ByVal lpOverlapped As LongPtr) As Long
    Private Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
    Private Declare PtrSafe Function WriteFile Lib "kernel32" (ByVal hFile As LongPtr, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long
    Private Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal hFile As LongPtr, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As LongPtr) As Long
    Private Declare PtrSafe Function DisconnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As LongPtr) As Long
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
#Else
    ' 32bit用は省略（必要ならLongPtrをLongに）
#End If

Private Const PIPE_ACCESS_DUPLEX As Long = &H3 ' 送受信可能
Private Const PIPE_TYPE_BYTE As Long = &H0     ' バイトモード
Private Const PIPE_WAIT As Long = &H0          ' ブロッキングモード（手動実行の要！）
Private Const INVALID_HANDLE_VALUE As LongPtr = -1
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const OPEN_EXISTING As Long = 3

' 共有のパイプ名
Private Const PIPE_NAME As String = "\\.\pipe\MyVbaPADHost"


Private hPipe As LongPtr

' --- 1. Excelに接続して命令を送る ---
Sub Step1_ConnectAndSend()
    ' Excelが開設したパイプに接続
    hPipe = CreateFile(PIPE_NAME, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)
    
    If hPipe = INVALID_HANDLE_VALUE Then
        MsgBox "Excelのパイプが見つかりません。Step0を実行しましたか？", vbCritical
        Exit Sub
    End If
    Debug.Print "Excelのパイプに接続成功！"
    
    ' Excelに命令を送る
    Dim commandMsg As String
    commandMsg = "URL遷移をお願いします！"
    
    Dim buffer() As Byte
    Dim bytesWritten As Long
    buffer = StrConv(commandMsg, vbFromUnicode)
    
    WriteFile hPipe, buffer(0), UBound(buffer) + 1, bytesWritten, 0
    Debug.Print "Excelに命令を送信しました！"
End Sub

' --- 5. Excelからの結果を受け取る ---
Sub Step5_ReceiveFromExcel()
    Dim buffer(0 To 1023) As Byte
    Dim bytesRead As Long
    Dim resultMsg As String
    
    ' Excelからの返信を待つ
    ReadFile hPipe, buffer(0), 1024, bytesRead, 0
    
    If bytesRead > 0 Then
        resultMsg = StrConv(LeftB(buffer, bytesRead), vbUnicode)
        MsgBox "Excelからの報告: " & vbCrLf & resultMsg, vbInformation, "ミッション完了"
    End If
    
    ' パイプをお片付け
    CloseHandle hPipe
End Sub
