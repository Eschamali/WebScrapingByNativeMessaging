Attribute VB_Name = "DemoPAD"
Option Explicit


Sub DemoPAD()
    Dim PADTest As New PADBrowser
    Debug.Print "初期化結果：" & PADTest.start(isDebugEnabled:=True)


    Dim ResultPAD As Object
    Set ResultPAD = PADTest.invokeMethod("GetAllTabsRequest")
    Stop



End Sub
