VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PADBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'     ブラウザ拡張機能の「ネイティブ メッセージング」機能を実際に利用する際のやり取りを行います
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                               ■■■ 必要なWindowsAPI宣言 ■■■
'***************************************************************************************************
Private Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" (ByVal hwnd As LongPtr, ByVal pszPath As String, ByVal psa As LongPtr) As Long



'***************************************************************************************************
'                                  ■■■ 関連の定数定義 ■■■
'***************************************************************************************************
Private Const ThisClassName As String = "PADBrowser"                'ログのトレース用

Private Const ERR_TIMEOUT   As Long = 900
Private Const ERR_PIPE      As Long = 901
Private Const ERR_PROTOCOL  As Long = 902



'***************************************************************************************************
'                                  ■■■ 関連の共通変数定義 ■■■
'***************************************************************************************************
Private core        As PADCore

Private ErrorDetail As New WinApiError  'エラーコードから、詳細を取得するやつ

Private LogControl      As New Logger        'ログレベルの制御
Private Utf8Converter   As New CharacterCodeConversion
Private JsonConverter   As New WebJsonConverter
Private timerOut        As Long
Private lngLastID       As Long
Private brID            As String
Private doLog As Boolean                'ログファイルにも、デバッグメッセージを残すか
Private logPath As String               'ログファイルの保存フォルダパス
Private seeRawSendMsgDbg As Boolean     'set to True if needing to see the raw json msg received from the browser for debugging
Private objAllMessages As Dictionary    '拡張機能からのイベントデータを蓄積するための変数です



'***************************************************************************************************
'                                  ■■■ メイン処理 ■■■
'***************************************************************************************************
'* 機能　　：Method名と、parametersを渡して、低レベルにてブラウザ制御をします
'---------------------------------------------------------------------------------------------------
'* 返り値　：拡張機能へのコマンド実行結果のうち、`Result`オブジェクトのみ。失敗時は、`Nothing`となります。
'* 引数　　：methodName     name項目にあたります
'            params         parameters項目にあたります。Json-Dictionary形式で、KeyとItemを事前に定義したオブジェクト変数を指定します。
'            StopError      `True`で、実行失敗時に`Err.Raise`で停止し、以降の処理を止めます
'***************************************************************************************************
Public Function invokeMethod(methodName As String, Optional params As Scripting.Dictionary, Optional StopError As Boolean = True) As Object
    Const FromProcedureName As String = ThisClassName & ".invokeMethod"


    Dim paramStr As String
    Dim results As Dictionary

    If params Is Nothing Then paramStr = "{}" Else paramStr = JsonConverter.ConvertToJson(params)
    Set results = SendMessage("{""name"":""" & methodName & """,""parameters"":" & paramStr & "}")

    'コマンド結果のうち、「result」項目オブジェクト飲み取り出して返却します
    If results.Exists("result") Then Set invokeMethod = results.Item("result")
    printMsg Debug_, "Invoked """ & methodName & """ with result: " & JsonConverter.ConvertToJson(results), FromProcedureName
End Function

'***************************************************************************************************
'* 機能　　：標準入出力の設定情報のみ反映します
'---------------------------------------------------------------------------------------------------
'* 注意事項：既に、アクティベーション処理を行った時に使うこと
'***************************************************************************************************
Public Sub reattach()
    Set core = New PADCore
End Sub



'***************************************************************************************************
'                                  ■■■ 核となる部分 ■■■
'***************************************************************************************************
'* 機能　　：CDP にメッセージを送信し、Chromium ブラウザでメッセージが受信されているかどうかを確認します
'---------------------------------------------------------------------------------------------------
'* 返り値　：拡張機能へのコマンド の実行結果そのまんま(Json-Dictionary 形式)
'* 引数　　：strMessage             Chromium ブラウザに送る拡張機能へのコマンド(プレーンテキスト)
'            InternalUseMessages    非同期イベントをクラス内部で後程、使用したいときに、`New Dictionry`として、渡して有効化します。
'***************************************************************************************************
Private Function SendMessage(strMessage As String, Optional InternalUseMessages As Scripting.Dictionary) As Dictionary
    Const FromProcedureName As String = ThisClassName & ".SendMessage"


    Dim strRes As String
    Dim intRes As Long

    Dim lngCurrentId As Long
    lngCurrentId = lngLastID
    lngLastID = lngLastID + 1

    Dim ErrorCode_PeekNamedPipe As Long, ErrorCode_ReadFile As Long


    '送信前にある貯まったデータを破棄する
    If objAllMessages Is Nothing And InternalUseMessages Is Nothing Then
        printMsg Trace_, "Starts discarding browser event information...", FromProcedureName

        '受信データ0が来るまで、ループ受信破棄を行います
        Do
            '受信処理
            intRes = core.readProcSTD(strRes, ErrorCode_PeekNamedPipe, ErrorCode_ReadFile)
            printMsg Trace_, "Discard - ResSize = " & intRes, FromProcedureName

            '返り値チェック
            If intRes > 0 Then
               'データありだが、あくまで破棄処理のため何もしない(強いていうなら、破棄するメッセージをそのまま出す)
               printMsg Trace_, "Emptying the message buffer...", FromProcedureName
               If seeRawSendMsgDbg Then printMsg Trace_, "Discarded - strRes = " & strRes, FromProcedureName

            Else
                'データなしの場合は、エラーコードチェックを行う
                If ErrorCode_PeekNamedPipe Then
                    'PeekNamedPipe にてエラーがあった場合
                    printMsg ERROR_, "Error PeekNamedPipe in readProcCDP. ErrorCode:" & ErrorCode_PeekNamedPipe & " - " & ErrorDetail.GetMessage(ErrorCode_PeekNamedPipe, "kernel32"), FromProcedureName, doRaiseError:=True

                ElseIf ErrorCode_ReadFile Then
                    'ReadFile にてエラーがあった場合
                    printMsg ERROR_, "Error ReadFile in readProcCDP. ErrorCode:" & ErrorCode_ReadFile & " - " & ErrorDetail.GetMessage(ErrorCode_ReadFile, "kernel32"), FromProcedureName, doRaiseError:=True

                Else
                    'データがまだ、ない。つまり、破棄完了
                    printMsg Trace_, "Message queue destruction completed.", FromProcedureName
                End If
            End If
        Loop Until intRes < 1
    End If


    '--------------------------------------------------
    '拡張機能へコマンド送信するフェーズ
    '--------------------------------------------------
    strMessage = Left(strMessage, Len(strMessage) - 1)
    strMessage = strMessage & ", ""requestId"":""" & lngCurrentId & """}"

    printMsg Debug_, "sendPAD-JsonCommand = " & strMessage, FromProcedureName
    core.writeProc strMessage


    '------------------------------------------------
    'ブラウザからの応答を確認するフェーズ
    '------------------------------------------------
    Dim timerStart As Double: timerStart = Timer

    '受信処理→Json処理→受信処理→Json処理→…　というロジックでループ処理を行います
    'CDP-Jsonコマンドを送った際のIDが返ってくるまで、ループ受信処理を行いますが、タイムアウトになったら、抜けます
    Dim ResultSendCommand As Boolean                                                             'We read from edge and process messages until we receive a message with our ID
    printMsg Trace_, "Starting Processes incoming messages from the browser...", FromProcedureName
    Do
        '受信処理
        intRes = core.readProcSTD(strRes, ErrorCode_PeekNamedPipe, ErrorCode_ReadFile)
        printMsg Trace_, "ResSize = " & intRes, FromProcedureName

        '返り値チェック
        If intRes > 0 Then
            'データあり。1件ずつとして保証されてるはず
            If seeRawSendMsgDbg Then printMsg Trace_, "Raw - strRes = " & strRes, FromProcedureName

            '受信データのチェックを行うプロシージャへ遷移します。
            If BrowserReceivedDataCheck(strRes, InternalUseMessages, lngCurrentId) Then
                '直近に送ったコマンド結果であることを確認できたので、受信ループ終了フラグを立てて、コマンド結果も返す
                ResultSendCommand = True
                Set SendMessage = JsonConverter.ParseJson(strRes)    'コマンド結果を返却
            End If
        Else
            'データなしの場合は、エラーコードチェックを行う
            If ErrorCode_PeekNamedPipe Then
                'PeekNamedPipe にてエラーがあった場合
                printMsg ERROR_, "Error PeekNamedPipe in readProcSTD. ErrorCode:" & ErrorCode_PeekNamedPipe & " - " & ErrorDetail.GetMessage(ErrorCode_PeekNamedPipe, "kernel32"), FromProcedureName, doRaiseError:=True
            ElseIf ErrorCode_ReadFile Then
                'ReadFile にてエラーがあった場合
                printMsg ERROR_, "Error ReadFile in readProcSTD. ErrorCode:" & ErrorCode_ReadFile & " - " & ErrorDetail.GetMessage(ErrorCode_ReadFile, "kernel32"), FromProcedureName, doRaiseError:=True
            Else
                'データがまだ、ない
                printMsg Trace_, "I used `readProcSTD` but it returned 0 Bytes.", FromProcedureName
            End If

            DoEvents
        End If

        'いくつかのログ表示
        If Not (ResultSendCommand) Then
            'コマンド結果が、まだ来ない
            printMsg Trace_, "The command result from readProcCDP has not yet been received, so reception continues.", FromProcedureName
        End If

        '一定時間経ってもコマンド結果が来ないなら、ループから抜けます
        If Timer - timerStart > timerOut Then Err.Raise ERR_TIMEOUT, FromProcedureName, "Timeout waiting for browser to return message." Else printMsg Trace_, Timer - timerStart & " seconds have passed...", FromProcedureName
    Loop Until ResultSendCommand
End Function

'***************************************************************************************************
'* 機能　　：ブラウザからの受信データの構造をチェックします
'---------------------------------------------------------------------------------------------------
'* 返値    ：TRUE :受信ループを止める合図。
'               ・`returnPADJson`に、ExcelからのPAD-Jsonコマンド結果がある
'               ・コマンド構文エラーによる、ID項目なし
'
'            FALSE:受信ループを継続する合図。
'               ・`JsonString`は、ブラウザからのイベント情報である
'
'* 引数　　：returnPADJson          1件分のJson-Dictionary形式のブラウザからのイベントやPAD-Jsonコマンド結果
'            InternalUseMessages    クラス内プロシージャで、イベント内容を扱う用として蓄積させる参照変数
'            checkCommandID         確認するコマンドID
'---------------------------------------------------------------------------------------------------
'* 詳細説明：イベントデータ or CDP-Jsonコマンド結果 or 構文エラー の確認をし、ブラウザイベントの場合はついでに、イベント情報を格納します。
'* 注意事項；CDP-Jsonコマンド結果IDの不一致(シングルスレッド配下なら多分ないはず...)の場合は、不整合により不安定につながるため、エラーで停止します
'***************************************************************************************************
Private Function BrowserReceivedDataCheck(returnPADJson As String, InternalUseMessages As Dictionary, checkCommandID As Long) As Boolean
    Const FromProcedureName As String = ThisClassName & ".BrowserReceivedDataCheck"


    '======================================= チェック開始 =======================================
    Dim ResultCommandID As Long, AsynchronousEventName As String
    Dim JsonDictionaryObj As Dictionary
    Dim FlagValue As Long

    Const resultCommand As Long = 2 ^ 0 'クラス内変数に記録してるIDのコマンド結果を検知出来た
    Const mismatchID    As Long = 2 ^ 1 'クラス内変数に記録してるIDとは違うコマンド結果が検知された。※Err.Raiseで止めるのであまり意味ないかも
    Const browserEvent  As Long = 2 ^ 2 '非同期イベントである
    Const EventCapture  As Long = 2 ^ 3 '非同期イベントキャプチャモードである
    Dim mask As Long: mask = resultCommand + mismatchID + browserEvent + EventCapture


    '--------- 1. 拡張機能からのイベントか？ ---------
    If GetContentsNotify(returnPADJson, AsynchronousEventName) Then
        FlagValue = FlagValue Or browserEvent

        '--------- 2. いずれかのDictionary変数が、非同期イベントキャプチャモードか？ ---------
        If (objAllMessages Is Nothing) And (InternalUseMessages Is Nothing) Then
            '両方とも`Nothing`なら、わざわざ中身を解析する意味がないので、パーサーは使わない
            printMsg Trace_, "The current process has asynchronous event capture mode disabled. Performance is prioritized, so the JSON parser is not used.", FromProcedureName
        Else
            'イベントキャプチャモードなので、JsonPurserを使用する
            FlagValue = FlagValue Or EventCapture
            printMsg Trace_, "A JSON parser is used for processing that requires asynchronous event information.", FromProcedureName
            Set JsonDictionaryObj = JsonConverter.ParseJson(returnPADJson)
        End If
    Else
        FlagValue = FlagValue Or resultCommand

        '--------- 2. `requestId`項目があるか？ ---------
        If GetContentsResult(returnPADJson, ResultCommandID) Then
            '--------- 3. そのrequestIdは、クラス内変数に記録してるrequestIdと一致してるか？ ---------
            If ResultCommandID = checkCommandID Then
                '正しいコマンド結果が返ったので次項に、受信ループ終了フラグを返すようにする
                printMsg Trace_, "It was confirmed that the execution command ID matches the one recorded in Excel. ID:" & checkCommandID, FromProcedureName
            Else
                '何かの拍子でずれてるようだ... シングルスレッド配下なら多分ないはず... だがずれてたら、不整合につながるため、エラーで止める
                FlagValue = FlagValue Or mismatchID
                printMsg ERROR_, "The command result was received(id:" & ResultCommandID & "), but it does not match the requestId(" & checkCommandID & ") recorded in Excel. Processing will be stopped due to the inconsistency.", FromProcedureName, doRaiseError:=True
            End If
        Else
            '`notify`も`requestId`項目もない場合、コマンドの構文エラーとして返った可能性があるとし、受信ループ終了とする
            printMsg WARN_, "Both the `requestId` and `notify` fields are missing. This may be a syntax error.", FromProcedureName
        End If
    End If


    '======================================= このプロシージャから抜けるか判定 =======================================
    Dim judgementBit As Long

    '1. 送ったコマンドの結果は来たので、受信ループを止めて、抜けるか？
    judgementBit = resultCommand
    If (FlagValue And mask) = judgementBit Then BrowserReceivedDataCheck = True: Exit Function

    '2. 受信ループは止めないが、イベントキャプチャしないモード(格納用のDictionaryがどちらもNothing or 古いSessionID)なので、次項の格納処理はせず、抜けるか？
    judgementBit = browserEvent
    If (FlagValue And mask) = judgementBit Then BrowserReceivedDataCheck = False: Exit Function

End Function

'***************************************************************************************************
'* 機能　　：ブラウザからの受信データの構造をチェックする際の1つ、`result`チェックを行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：存在有無論理値
'* 引数　　：JsonString         ネイティブ メッセージング プロトコルのレスポンス(Json文字列そのまんま)
'            idValueContents    コマンド結果のID値を格納する変数(参照渡しによって返します)
'***************************************************************************************************
Private Function GetContentsResult(jsonString As String, ByRef requestIdValueContents As Long) As Boolean
    Const FromProcedureName As String = ThisClassName & ".GetContentsID"


    '検索設定
    Const SearchKeyName     As String = "result"    '検索文字列
    Const SearchKeyName2    As String = "requestId" '検索文字列
    Const StartSearchPos    As Long = 3             'あるべき位置


    '1. 文字列を探索
    Dim StartPos As Long
    StartPos = InStr(StartSearchPos, jsonString, SearchKeyName)

    '2. 所定の位置にない場合は、存在しないとし、抜ける
    '下記の例に基づき、3文字目にない場合は、コマンド結果ではないとして、扱います。
    '・{"result":..."requestId":"0"}
    If StartPos = StartSearchPos Then
        '3. 後ろから探してみる
        StartPos = InStrRev(jsonString, SearchKeyName2) + Len(SearchKeyName2) + Len(""":""") '「requestId":"」分、ずらすイメージ

        '4. ID値の終点を求める
        Dim EndPos As Long
        EndPos = InStr(StartPos, jsonString, """,")

        '5. ID値を取り出す
        On Error Resume Next
            requestIdValueContents = Mid(jsonString, StartPos, EndPos - StartPos)
            If Err.Number Then
                GetContentsResult = False
                printMsg Trace_, "`requestId`という文字列はありますが、数字ではないようです...", FromProcedureName
            Else
                GetContentsResult = True
                printMsg Trace_, "This response is the command result. requestId:" & requestIdValueContents, FromProcedureName
            End If
        On Error GoTo 0
    Else
        GetContentsResult = False
        printMsg Trace_, "This response does not have an `requestId`.", FromProcedureName
    End If
End Function

'***************************************************************************************************
'* 機能　　：ブラウザからの受信データの構造をチェックする際の1つ、`notify`チェックを行います
'---------------------------------------------------------------------------------------------------
'* 返り値  ：存在有無論理値
'* 引数　　：JsonString             ネイティブ メッセージング プロトコルのレスポンス(Json文字列そのまんま)
'            methodValueContents    非同期イベントのイベント名(参照渡しによって返します)
'***************************************************************************************************
Private Function GetContentsNotify(jsonString As String, ByRef notifyValueContents As String) As Boolean
    Const FromProcedureName As String = ThisClassName & ".GetContentsMethod"


    '検索設定
    Const SearchKeyName     As String = "notify"    '検索文字列
    Const StartSearchPos    As Long = 3             'あるべき位置


    '1. 文字列を探索
    Dim StartPos As Long
    StartPos = InStr(StartSearchPos, jsonString, SearchKeyName)

    '2. 所定の位置にない場合は、存在しないとし、抜ける
    '下記の例に基づき、3文字目にない場合は、非同期イベントではないとして、扱います。
    '・{"notify":"...
    If StartPos = StartSearchPos Then
        GetContentsNotify = True

        '3. 開始位置をうまい具合にずらす
        StartPos = StartSearchPos + Len(SearchKeyName) + Len(""":""") '「notify":"」分、ずらすイメージ

        '4. イベント名の終点を求める
        Dim EndPos As Long
        EndPos = InStr(StartPos, jsonString, """,")

        '5. イベント名を取り出す
        notifyValueContents = Mid(jsonString, StartPos, EndPos - StartPos)
        printMsg Trace_, "This response is an asynchronous event: " & notifyValueContents, FromProcedureName
    Else
        GetContentsNotify = False
        printMsg Trace_, "This response does not have an `method`.", FromProcedureName
    End If
End Function

'***************************************************************************************************
'* 機能　　：ログメッセージを綺麗化します
'---------------------------------------------------------------------------------------------------
'* 引数　　：LogLevel_      ログレベル
'            strMsg         ログメッセージ内容
'            From           このメソッドを呼び出すプロシージャ名
'            isHeader
'            doRaiseError   `True`でVBA処理停止とします
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean, Optional doRaiseError As Boolean)
    Dim logName As String, strFormattedMsg As String

    If isHeader Then strFormattedMsg = String(100, "-") & vbNewLine & strMsg & vbNewLine & String(100, "-") _
    Else strFormattedMsg = " | " & brID & " | " & strMsg

   'Perform msg logging if needed
    If doLog Then
        logName = "log" & UCase(Format(Now, "ddMMMyy")) & ".pad.log"
        If Dir(logPath, vbDirectory) = "" Then MkDir logPath
    End If

   'Perform msg output to Immediate Window if needed
    With LogControl
        Select Case LogLevel_
            Case LogLevelName.Trace_: .LogTrace strFormattedMsg, From, logPath, logName
            Case LogLevelName.Debug_: .LogDebug strFormattedMsg, From, logPath, logName
            Case LogLevelName.info_: .LogInfo strFormattedMsg, From, logPath, logName
            Case LogLevelName.WARN_: .LogWarn strFormattedMsg, From, logPath, logName
            Case LogLevelName.ERROR_: .LogError strFormattedMsg, From, logPath, logName, Err.Number
            Case Else: 'ログ出力無効化
        End Select
    End With


    'ブラウザ自身や、パイプ通信自体に問題や、タイムアウトになったと判断された場合は、このツールでの処理(CDP-Json)ができないので、エラーで停止させておく
    If doRaiseError Then Err.Raise ERR_PROTOCOL, From, description:=strMsg

End Sub



'***************************************************************************************************
'                                  ■■■ 初期設定関連処理 ■■■
'***************************************************************************************************
'* 機能　　：このExcelを「ネイティブ メッセージング ホスト」にするための初期設定を行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：成功可否論理値
'---------------------------------------------------------------------------------------------------
'* 機能説明：「https://developer.chrome.com/docs/extensions/develop/concepts/native-messaging?hl=ja」に従ったレジストリ登録をして、
'            ブラウザからへExcelにパイプ接続する下準備を行います。
'* 注意事項：これを行うと、PowerAutomateDeskTopからの操作ができなくなります
'***************************************************************************************************
Public Function InitialSetup() As Boolean
    'ファイル名設定など
    Const RapperBatName As String = "StartNewInstanceExcel.bat"
    Const SaveJsonName As String = "host_manifest.json"


    '1. セルから設定を取得
    With ShSetting01_RegistrySet
        '1-1. 拡張機能情報
        Dim extID As String: extID = .Range(.UseRangeName(1, "PADCore.InitRegistry")).value
        Dim appName As String: appName = .Range(.UseRangeName(2, "PADCore.InitRegistry")).value
        Dim description As String: description = .Range(.UseRangeName(3, "PADCore.InitRegistry")).value

        '1-2. ネイティブ メッセージング ホストの構成情報の保存パスの指定
        Dim FolderPath As String: FolderPath = .Range(.UseRangeName(4, "PADCore.InitRegistry")).value

        '1-3. 使用ブラウザ
        '※一旦は、Edge or Chrome のみとする
        Dim ChromeUse As Boolean: ChromeUse = .Range(.UseRangeName(5, "PADCore.InitRegistry")).value
    End With

    '2. フォルダがなければ作成
    Dim ResultCode As Long
    ResultCode = SHCreateDirectoryEx(0&, FolderPath, 0&)

    'エラーチェック
    Select Case ResultCode
        Case 0, 183
        Case Else: MsgBox "ネイティブ メッセージング ホストの構成情報の保存先のフォルダ作成に失敗しました。" & vbCrLf & vbCrLf & "＜原因＞" & vbCrLf & ErrorDetail.GetMessage(ResultCode), vbCritical, "ErrorCode：" & ResultCode: Exit Function
    End Select

    '3. ラッパーのBatファイルパスを取得し、JSON用にバックスラッシュを二重化
    Dim escapedPath As String: escapedPath = Replace(FolderPath & "\" & RapperBatName, "\", "\\")

    '4. ネイティブ メッセージング ホストの構成情報になるJSON文字列の生成
    Dim jsonContent As String
    jsonContent = "{" & vbCrLf & _
        "  ""name"": """ & appName & """," & vbCrLf & _
        "  ""description"": """ & description & """," & vbCrLf & _
        "  ""path"": """ & escapedPath & """," & vbCrLf & _
        "  ""type"": ""stdio""," & vbCrLf & _
        "  ""allowed_origins"": [""chrome-extension://" & extID & "/""]" & vbCrLf & _
        "}"

    '5. 4の内容をJSONファイルとして、書き出し
    With Utf8Converter
        .BytesToSaveFile .BytesFromString(jsonContent), FolderPath, SaveJsonName
    End With

    '6. ラッパーBatに記述する内容を生成
    Dim BatContent As String
    BatContent = "@echo off" & vbCrLf & Chr(34) & Application.Path & "\EXCEL.EXE"" /x """ & OneDrivePathToLocalPath(ThisWorkbook.FullName) & Chr(34)

    '7. 6の内容をBatファイルとして、書き出し
    Dim fNum As Integer: fNum = FreeFile
    Open FolderPath & "\" & RapperBatName For Output As #fNum
        Print #fNum, BatContent;
    Close #fNum

    '8. レジストリの書き換え準備 (WMI使用)
    Const HKCU              As Long = &H80000001                '現在のログインユーザー単位。本家のPowerAutomateDesktopもこれに倣ってるので
    Const TargetBaseKeyName As String = "NativeMessagingHosts"

    Dim objReg As Object
    Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

    '9. 対象レジストリキーパス生成
    '※一旦は、Edge or Chrome のみとする
    Dim TargetKeyPath As String
    TargetKeyPath = "Software\" & IIf(ChromeUse, _
                    "Google\Chrome\", _
                    "Microsoft\Edge\") & TargetBaseKeyName & "\" & appName

    '10. キーを作成
    '※既存でも0が返ります
    ResultCode = objReg.CreateKey(HKCU, TargetKeyPath)
    If ResultCode <> 0 Then MsgBox "レジストリキーパス「" & TargetKeyPath & "」の作成に失敗しました。", vbCritical, "ErrorCode：" & ResultCode: Exit Function

    '11. 値をセット (JSONファイルへのパス)
    '※既定値なので値の名前に空文字を指定)
    ResultCode = objReg.SetStringValue(HKCU, TargetKeyPath, "", FolderPath & "\" & SaveJsonName)
    If ResultCode <> 0 Then MsgBox "レジストリキー名「" & appName & "」の規定値への書き込みに失敗しました。", vbCritical, "ErrorCode：" & ResultCode: Exit Function

    '12. Done!
    InitialSetup = True
End Function

'***************************************************************************************************
'* 機能　　：ブラウザを操作する前のネゴシエーションを行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：成功可否論理値
'---------------------------------------------------------------------------------------------------
'* 機能説明：公式のPowerAutomate拡張機能は起動直後、「LoadScriptsRequest」という初期化命令しか受け付けない状態になっているため、それを送れるようにします
'* 注意事項：・よって必ずこのメソッドをまず1回以上呼び出す必要があります
'            ・引数は一旦全て、オプション化してます
'***************************************************************************************************
Public Function start(Optional backgroundVersion As Long = 4, Optional contentVersion As Long = 2, Optional apiVersion As Long = 4, Optional isDebugEnabled As Boolean) As Boolean
    Set core = New PADCore

    Dim initJson As String
    initJson = "{""name"":""LoadScriptsRequest"",""parameters"":{""backgroundVersion"":" & backgroundVersion & _
                                                                    ",""contentVersion"":" & contentVersion & _
                                                                    ",""apiVersion"":" & apiVersion & _
                                                                    ",""isDebugEnabled"":" & LCase(CStr(isDebugEnabled)) & "}}"

    '初期化コマンドを送ります
    Dim ResultCmd As Dictionary: Set ResultCmd = SendMessage(initJson)

    '結果を確認
    start = CBool(ResultCmd("result"))
End Function

'***************************************************************************************************
'* 機能　　：このExcelが、OneDrive上で実行されてる場合のパス変換処理を行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：ローカルパス
'* 引数　　：Path                   基本は、`thisworkbook.path`を指定
'            UsePrivateOneDrive     社内個人OneDriveの場合は、`False`にしてください
'---------------------------------------------------------------------------------------------------
'* 機能説明：開いてるExcelがOneDriveにあると、`thisworkbook.path`がインターネット上のURLになってしまい、一部操作ができなくなる問題に対処した物となります。
'            純ローカルなら、そのまま返します。
'            個人向けOneDrive と ビジネス向け個人OneDrive に対応してます。先頭の定数で、スイッチングしてください
'
'* 注意事項：SharePointの場合は、自力でコードを書く必要があります
'***************************************************************************************************
Private Function OneDrivePathToLocalPath(Path As String, Optional UsePrivateOneDrive As Boolean = True) As String
    'http始まりじゃないなら、そのまま返して終了
    If Left(Path, 4) <> "http" Then OneDrivePathToLocalPath = Path: Exit Function

    '個人OneDriveモードなら識別番号分、ローカルパスに置き換えて結合
    If UsePrivateOneDrive Then
        OneDrivePathToLocalPath = Environ("OneDrive") & Mid(Path, 41)
    
    '個人BusinessOneDriveモードなら"Documents"以降のパスを抜き出して、ローカルパスと結合
    Else
        OneDrivePathToLocalPath = Environ("OneDriveCommercial") & Evaluate("TEXTAFTER(""" & Path & """,""/Documents"")")
    End If
End Function



Private Sub Class_Initialize()
    'Assign an objID for easy debug identification
    brID = "BRID" & Format(Rnd * 1000, "000")

    '------------------------- クラス内変数の一部を初期化 -------------------------
    '設定シートの各セルから設定値を取得し、クラス内変数に適用します
    With ShSetting01_RegistrySet
        '1. イミディエイトに関する設定を行う
        seeRawSendMsgDbg = .Range(.UseRangeName(8, "CDPBrowser.Class_Initialize")).value           'ブラウザから受信した生のJSONメッセージをイミディエイトに表示させるか？

        '2. パスが設定されてたら、ログファイル保存を有効化
        Dim LogFolderPath As String: LogFolderPath = .Range(.UseRangeName(9, "CDPBrowser.Class_Initialize")).value
        If LogFolderPath <> "" Then
            doLog = True                'ログファイルとして保存するか？
            logPath = LogFolderPath     '保存先のフォルダパスを設定
        End If

    End With

    'タイムアウト設定(デフォルトで、45s)
    timerOut = 45
    '------------------------------------------------------------------------------
End Sub
