Attribute VB_Name = "DemoPAD"
'***************************************************************************************************
'                ブラウザ拡張機能の「PowerAutomate」機能を模倣するDemoコードです
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                                       ■■■ はじめに ■■■
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
'                                       ■■■ 基本 ■■■
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

    Dim ResultPAD As Variant
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

    Dim ResultPAD As Variant
    Set ResultPAD = PADTest.invokeName("GetAllTabsRequest")

    'イベント情報をDownloadsフォルダに保存
    With UTF8Conv
        .BytesToSaveFile .BytesFromString(JsonDicObj.ConvertToJson(PADTest.BrowserEvents)), Environ("UserProfile") & "\Downloads", "Event.json"
    End With
End Sub



'***************************************************************************************************
'                                       ■■■ 実践 ■■■
'***************************************************************************************************
Sub Webページ遷移()
    '対象のURLを指定してするか、下記のページを用意してください
    Const TargetURL As String = "https://news.google.com/home?hl=ja&gl=JP&ceid=JP:ja"


    '1. アクティベーション処理を飛ばして、設定
    Dim PADTest As New PADBrowser
    PADTest.reattach

    '2. 現在開いてるタブ情報一式を取得
    Dim ResultPAD As Variant
    Set ResultPAD = PADTest.invokeName("GetAllTabsRequest")

    '3. 対象URLを見つけたら、`tabid`を取得する
    '※同じURLがあっても、最初に見つかったのを対象とします
    Dim TargetTabID As Long
    Dim i As Long, tmp
    For Each tmp In ResultPAD
        If tmp("url") = TargetURL Then TargetTabID = tmp("id"): Exit For
    Next

    '4. URL遷移を行う
    Dim PADparameter As Dictionary
    Set PADparameter = New Dictionary
    PADparameter.Add "tabId", TargetTabID
    PADparameter.Add "url", "https://developer.chrome.com/docs/extensions/develop/concepts/native-messaging?hl=ja"
    Debug.Print "ページ遷移結果：" & PADTest.invokeName("NavigateToUrlRequest", PADparameter)
End Sub

Sub 最後のアクティブウィンドウに対してJavaScript関数を実行()
    'アクティブウィンドウ関係のイベント名
    Const ActiveWindowEventName As String = "window_focused"


    '1. アクティベーション処理を飛ばして、設定
    Dim PADTest As New PADBrowser
    PADTest.reattach

    '2. イベントキャプチャしながら、アクティブウィンドウを特定
    Set PADTest.BrowserEvents = New Dictionary
    Dim ResultPAD As Variant, WindowCount As Long
    Set ResultPAD = PADTest.invokeName("GetAllWindowsRequest")
    WindowCount = ResultPAD.Count

    '1つしかない場合は実質それが最後のアクティブウィンドウなのでイベントキャプチャ内容の確認は飛ばす
    Dim ActiveWindowTabs As Dictionary
    If WindowCount > 1 Then
        '3. アクティブウィンドウ情報を取得
        Dim ActiveWindowsInfo As Collection, focusedCount As Long
        If Not (PADTest.BrowserEvents("EventNotify").Exists(ActiveWindowEventName)) Then MsgBox "ブラウザから、アクティブウィンドウ情報の取得に失敗しました。" & vbCrLf & "VBAの制約上、各ブラウザウィンドウを最低1回ずつフォーカスする必要があります。", vbCritical, "アクティブウィンドウIDの取得に失敗": Exit Sub
        Set ActiveWindowsInfo = PADTest.BrowserEvents("EventNotify")(ActiveWindowEventName)
        focusedCount = ActiveWindowsInfo.Count

        '4. 最後のアクティブウィンドウIDを取得
        Dim LastWindowID As Long
        LastWindowID = ActiveWindowsInfo(focusedCount)("windowId")

        '5. アクティブウィンドウコレクションを取得
        Dim tmp
        For Each tmp In ResultPAD
            If tmp("id") = LastWindowID Then Set ActiveWindowTabs = tmp: Exit For
        Next
    Else
        '先頭のウィンドウ情報としてセット
        Set ActiveWindowTabs = ResultPAD(1)
    End If

    '6. アクティブなタブを探します
    Dim ActiveTabID As Long
    For Each tmp In ActiveWindowTabs("tabs")
        If tmp("active") Then ActiveTabID = tmp("id"): Exit For
    Next

    '7. JavaScript関数を実行します
    Dim PADparameter As Dictionary
    Set PADparameter = New Dictionary
    PADparameter.Add "tabId", ActiveTabID
    PADparameter.Add "code", "alert('VBAからの電撃訪問です！" & WorksheetFunction.Unichar(129760) & "');"
    Debug.Print "実行結果：" & PADTest.invokeName("RunJavaScriptRequest", PADparameter)
End Sub
