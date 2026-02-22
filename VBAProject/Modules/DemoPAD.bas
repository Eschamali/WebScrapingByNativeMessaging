Attribute VB_Name = "DemoPAD"
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
'* 機能　　：現在開いているタブ情報を取得します
'---------------------------------------------------------------------------------------------------
'* 注意事項：・ログレベル：Debug にすると分かりやすいと思います
'            ・アクティベーション処理が既に済んでるとします
'***************************************************************************************************
Sub ShowAllTabs()
    Dim PADTest As New PADBrowser
    PADTest.reattach

    Dim ResultPAD As Object
    Set ResultPAD = PADTest.invokeMethod("GetAllTabsRequest")
End Sub
