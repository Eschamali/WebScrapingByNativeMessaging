Attribute VB_Name = "DemoConnectPipeWithName"
'***************************************************************************************************
'        PowerAutomateDesktop と PAD.BrowserNativeMessageHost.exe のやり取りの核と言える
'　「名前付きパイプ（サーバー化）」による双方向なやり取りの概念実証デモンストレーションです
'
'                           このコードは、ExcelVBA(サーバー)用です
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

' --- 0. パイプ開設と接続待ち ---
Sub Step0_OpenServer()
    ' パイプを作成
    hPipe = CreateNamedPipe(PIPE_NAME, PIPE_ACCESS_DUPLEX, PIPE_TYPE_BYTE Or PIPE_WAIT, 1, 1024, 1024, 0, 0)
    If hPipe = INVALID_HANDLE_VALUE Then
        MsgBox "パイプの作成に失敗しました??", vbCritical
        Exit Sub
    End If
    
    Debug.Print "パイプを開設しました。Wordからの接続を待っています..."
    
    ' ★注意★ ここでExcelはWordが繋いでくるまで「フリーズ（待機状態）」になります！
    ConnectNamedPipe hPipe, 0
    
    ' Wordが繋ぐとフリーズが解けてここに進む
    Debug.Print "Wordが接続してきました！"
End Sub

' --- 2. Wordからの命令を受信 ---
Sub Step2_ReceiveFromWord()
    Dim buffer(0 To 1023) As Byte
    Dim bytesRead As Long
    Dim msg As String
    
    ' Wordからのメッセージを読み取る（文字が来るまで待機）
    ReadFile hPipe, buffer(0), 1024, bytesRead, 0
    
    If bytesRead > 0 Then
        ' バイト配列を文字列に変換 (UTF-8等ではなく手抜きでShift-JIS変換)
        msg = StrConv(LeftB(buffer, bytesRead), vbUnicode)
        Debug.Print "Wordからの命令: " & msg
        ' セルに書き出してもOK！
    End If
End Sub

' --- 3. ブラウザへ送信 ＆ 4. Wordへ結果を返す ---
Sub Step3and4_SendResultToWord()
    ' 【本来のStep3】 ここで受け取った命令(msg)をもとに Select Case 等で
    ' ネイティブメッセージングの SendMessage / ReceiveMessage を行います。
    ' 今回は概念実証なので、モック（ダミー結果）を作ります。
    
    Dim resultMsg As String
    resultMsg = "ブラウザ操作完了！(Excelより愛を込めて)"
    
    Dim buffer() As Byte
    Dim bytesWritten As Long
    buffer = StrConv(resultMsg, vbFromUnicode)
    
    ' Wordへ結果を送信！
    WriteFile hPipe, buffer(0), UBound(buffer) + 1, bytesWritten, 0
    Debug.Print "Wordへ結果を返しました！"
    
    ' 最後にパイプをお片付け
    ' ※Wordで受信後、続行すること
    Stop
    DisconnectNamedPipe hPipe
    CloseHandle hPipe
End Sub
