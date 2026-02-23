Attribute VB_Name = "DemoPAD"
'***************************************************************************************************
'                ブラウザ拡張機能の「PowerAutomate」機能を模倣するDemoコードです
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'* 機能　　：このExcelBookをPowerAutomate拡張機能のネイティブメッセージングホストに書き換える禁断魔術を行います
'***************************************************************************************************
Sub 初期設定()
    Dim Initial As New PADBrowser
    Dim Result As Boolean: Result = Initial.InitialSetup

    '成功したら、通知
    If Result Then
        Dim MsgboxCode As Long
        MsgboxCode = MsgBox("このマクロブックを、PowerAutomate拡張機能のネイティブメッセージングホストに書き換えました。" & vbCrLf & "このBookを閉じて、拡張機能のインストールページに進んでもよろしいですか？", vbInformation + vbYesNo, "Done")

        '「はい」を選んだら、拡張機能のリンクを開き、このブックを閉じます
        If MsgboxCode = vbYes Then
            CreateObject("WScript.Shell").Run "https://microsoftedge.microsoft.com/addons/detail/microsoft-power-automate/kagpabjoboikccfdghpdlaaopmgpgfdc"
            ThisWorkbook.Close
            'Application.Quit   '同一インスタンスの開いてるブックも巻き込んで消えるので却下
        End If
    End If
End Sub

'***************************************************************************************************
'* 機能　　：PowerAutomate拡張機能 のアクティベーションだけを行います
'---------------------------------------------------------------------------------------------------
'* 機能説明：PowerAutomate拡張機能 の場合は初回に、アクティベーション処理的なコマンドが必要なのでこれを設けます
'* 注意事項：メインのコマンド(GetAllTabsRequestなど)を行う前に、これが必ず必要です
'***************************************************************************************************
Sub StartActivation()
    Dim Initial As New PADBrowser
    Debug.Print "アクティベーション結果：" & Initial.start(isDebugEnabled:=True)
End Sub

'***************************************************************************************************
'* 機能　　：現在開いているタブ情報を保存します
'---------------------------------------------------------------------------------------------------
'* 注意事項：アクティベーション処理が既に済んでるとします
'***************************************************************************************************
Sub ShowAllTabs()
    Dim PADTest As New PADBrowser
    Dim UTF8Conv As New CharacterCodeConversion
    Dim JsonDicObj As New WebJsonConverter

    'アクティベーション処理を飛ばして、設定
    PADTest.reattach

    Dim ResultPAD As Object
    Set ResultPAD = PADTest.invokeName("GetAllTabsRequest")

    'タブ情報をDownloadsフォルダに保存
    With UTF8Conv
        .BytesToSaveFile .BytesFromString(JsonDicObj.ConvertToJson(ResultPAD)), Environ("UserProfile") & "\Downloads", "GetAllTabsRequest.json"
    End With
End Sub

'***************************************************************************************************
'* 機能　　イベントキャプチャテスト
'---------------------------------------------------------------------------------------------------
'* 注意事項：・ブラウザのウィンドウアクティブ操作を済ませた後、実行してください
'            ・アクティベーション処理が既に済んでるとします
'***************************************************************************************************
Sub SaveEvent()
    Dim PADTest As New PADBrowser
    Dim UTF8Conv As New CharacterCodeConversion
    Dim JsonDicObj As New WebJsonConverter

    'アクティベーション処理を飛ばして、設定
    PADTest.reattach

    'イベントキャプチャを有効化
    Set PADTest.BrowserEvents = New Dictionary

    Dim ResultPAD As Object
    Set ResultPAD = PADTest.invokeName("GetAllTabsRequest")

    'イベント情報をDownloadsフォルダに保存
    With UTF8Conv
        .BytesToSaveFile .BytesFromString(JsonDicObj.ConvertToJson(PADTest.BrowserEvents)), Environ("UserProfile") & "\Downloads", "Event.json"
    End With
End Sub
