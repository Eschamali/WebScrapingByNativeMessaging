VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShSetting01_RegistrySet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit



'***************************************************************************************************
'                             ■■■ プロパティメソッド ■■■
'***************************************************************************************************
'* 機能　　：このシートのみ定義されてるセル名称を取得します
'---------------------------------------------------------------------------------------------------
'* 返り値　：IDに紐づくセル名称
'* 引数　　：RangeID        セルID
'            ErrorSource    この関数を呼び出したソース
'***************************************************************************************************
Property Get UseRangeName(RangeID As Long, ErrorSource As String) As String
    Select Case RangeID
        Case 1: UseRangeName = "拡張機能ID"
        Case 2: UseRangeName = "拡張機能内部名称"
        Case 3: UseRangeName = "説明文"
        Case 4: UseRangeName = "保存フォルダパス"
        Case 5: UseRangeName = "Chromeで使用"
        Case 6: UseRangeName = "LogLevel"
        Case 7: UseRangeName = "イミディエイトに表示"
        Case 8: UseRangeName = "ブラウザから受信した生のJSONをイミディエイトに表示"
        Case 9: UseRangeName = "ログファイル保存"
        Case Else: Err.Raise 1 + vbObjectError, ErrorSource, "渡されたセルIDは、定義されてません"
    End Select
End Property
