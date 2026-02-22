VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PADCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'       ブラウザ拡張機能の「ネイティブ メッセージング」機能を利用するためのコア部分となります
'       https://developer.chrome.com/docs/extensions/develop/concepts/native-messaging?hl=ja
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                           ■■■ やり取り等に必要なWindowsAPI宣言 ■■■
'***************************************************************************************************
Private Declare PtrSafe Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As LongPtr
Private Declare PtrSafe Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As LongPtr, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long
Private Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal hFile As LongPtr, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare PtrSafe Function WriteFile Lib "kernel32" (ByVal hFile As LongPtr, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long



'***************************************************************************************************
'                                   ■■■ 関連の定数定義 ■■■
'***************************************************************************************************
Private Const STD_INPUT_HANDLE  As Long = -10
Private Const STD_OUTPUT_HANDLE As Long = -11
Private Const STD_ERROR_HANDLE  As Long = -12



'***************************************************************************************************
'                                  ■■■ 関連の共通変数定義 ■■■
'***************************************************************************************************
Private Utf8Converter   As New CharacterCodeConversion
Private hStdIn          As LongPtr
Private hStdOut         As LongPtr
Private hStdErrOut      As LongPtr



'***************************************************************************************************
'                                  ■■■ メイン処理 ■■■
'***************************************************************************************************
'* 機能　　：拡張機能から送信された内容を受信します
'---------------------------------------------------------------------------------------------------
'* 返り値　：command結果のデータサイズ。
'* 引数　　：strData            拡張機能へ送ったコマンド結果。実質返り値です
'            Err_PeekNamedPipe  PeekNamedPipe にてエラーが起こった際のエラーコードを格納します。実質返り値です
'            Err_ReadFile       ReadFile にてエラーが起こった際のエラーコードを格納します。実質返り値です
'---------------------------------------------------------------------------------------------------
'* 機能説明：・読み取りは非ブロッキングです。データがなくても、次に進めれます。各種エラーコードは、第2引数以降で確認して下さい。
'            ・ネイティブ メッセージング プロトコル に従った方式で処理していきます
'***************************************************************************************************
Public Function readProcSTD(ByRef strData As String, ByRef Err_PeekNamedPipe As Long, ByRef Err_ReadFile As Long) As Long
    'エラーコードの初期化
    Err_PeekNamedPipe = 0
    Err_ReadFile = 0

    Dim lPeekData As Long
    Dim lngRes As Long

    '処理が成功すると、戻り値は 0 以外になります。
    lngRes = PeekNamedPipe(hStdIn, 0&, 0&, 0&, lPeekData, 0&)

    'API関数:PeekNamedPipe が失敗した場合は、0 を返すため、`Err_PeekNamedPipe`にエラーコードを格納し、抜ける
    If lngRes = 0 Then Err_PeekNamedPipe = Err.LastDllError: Exit Function


    'データがあったら2段階の操作を得て、データを取り出していきます
    If lPeekData > 0 Then
        '---------- 1. コマンド結果のサイズを取得 ----------
        '先頭4Byte(32bit)に、その情報が載ってるのでまずはそれを取り出します
        Const DataSize目印 As Long = 4
        Dim msgLen As Long
        Dim bytesRead As Long
        Dim bSuccess As Long
        bSuccess = ReadFile(hStdIn, msgLen, DataSize目印, bytesRead, ByVal 0&)

        'ちゃんと4Byteで取り出せたなら成功です
        If bytesRead = DataSize目印 Then
            '実際のデータサイズを返却
            readProcSTD = msgLen

            '---------- 2. 実際のコマンド結果を取得 ----------
            ' 2-1. 文字列を格納するバッファ（箱）を用意
            Dim buffer() As Byte
            ReDim buffer(msgLen - 1)

            ' 2-2. 本体のJSON文字列を読み取る
            bSuccess = ReadFile(hStdIn, buffer(0), msgLen, bytesRead, ByVal 0&)

            '0以外なら成功です
            If bSuccess Then
                ' 2-3. バイナリを文字列（UTF-8）に変換
                strData = Utf8Converter.BytesToString(buffer)
            Else
                'API関数:ReadFile が失敗した場合は、0 を返すため、`Err_ReadFile`にエラーコードを格納
                Err_ReadFile = Err.LastDllError
            End If
        Else
            '失敗なので、エラーコードを入れる
            Err_ReadFile = Err.LastDllError
        End If
    End If
End Function

'***************************************************************************************************
'* 機能　　：拡張機能へコマンドを送信します
'---------------------------------------------------------------------------------------------------
'* 返り値　：0が成功。それ以外は、エラーコードになります。
'* 引数　　：strData            コマンド内容
'            ErrorSendMode      `True` ：標準出力エラーとして送ります
'                               `False`：標準出力として送ります。コマンド構文に従う必要があります。
'---------------------------------------------------------------------------------------------------
'* 機能説明：ネイティブ メッセージング プロトコル に従った方式で送っていきます
'***************************************************************************************************
Public Function writeProc(ByVal strData As String, Optional ErrorSendMode As Boolean) As Long
    '1. 引数に応じて使うハンドル情報を切り替えます
    Dim UsePipeHandle As LongPtr: UsePipeHandle = IIf(ErrorSendMode, hStdErrOut, hStdOut)

    '2. VBAの「Unicode」文字列をUTF-8でエンコードしたバイト配列を取得し、メッセージの長さも取得します
    Dim BinData() As Byte: BinData = Utf8Converter.BytesFromString(strData)
    Dim msgLen As Long: msgLen = UBound(BinData) + 1

    '標準出力エラーの場合は下記の工程は飛ばせます
    Dim lngWritten As Long
    If Not (ErrorSendMode) Then
        '3. 最初に「メッセージの長さ」を4バイトで送る
        WriteFile UsePipeHandle, msgLen, 4, lngWritten, ByVal 0&
    End If

    '4. 次に、メッセージ本体のバイト配列として、コマンドを送ります
    writeProc = WriteFile(UsePipeHandle, ByVal VarPtr(BinData(0)), UBound(BinData) + 1, lngWritten, ByVal 0&)

    'エラー時は、エラーコードを返します
    If writeProc = 0 Then writeProc = Err.LastDllError
End Function



Private Sub Class_Initialize()
    'Excel自身の標準入出力とエラーパイプハンドル情報を取得します
    hStdIn = GetStdHandle(STD_INPUT_HANDLE)
    hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)
    hStdErrOut = GetStdHandle(STD_ERROR_HANDLE)
End Sub
