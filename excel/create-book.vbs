' ****************************
' Excel オブジェクト作成
' ****************************
Set App = CreateObject("Excel.Application")

' ****************************
' 警告を出さないようにする
' ****************************
App.DisplayAlerts = False

' ****************************
' ブック追加
' ****************************
App.Workbooks.Add()

' ****************************
' 追加したブックを取得
' ****************************
Set Workbook = App.Workbooks( App.Workbooks.Count )

' ****************************
' 現状、ブックにはシート一つ
' という前提で処理していますが
' 必要であれば、Book.Worksheets.Count
' で現在のシートの数を取得できます
' ****************************
Set Worksheet = Workbook.Worksheets( 1 )
Worksheet.Activate()

' ****************************
' Add では 第二引数に指定した
' オブジェクトのシートの直後に、
' 新しいシートを追加します。
' ****************************
Call Workbook.Worksheets.Add(,Worksheet)

' ****************************
' シート名設定
' ****************************
Workbook.Sheets(1).Name = "初期シート"
Workbook.Sheets(2).Name = "追加シート"

' ****************************
' データ操作
' ****************************
Workbook.Sheets(1).Activate()
Workbook.Sheets(1).Cells(1, 2) = "社員コード"
Workbook.Sheets(1).Range("B2") = "0001"

Workbook.Sheets(1).Range("B2").Select()
' https://docs.microsoft.com/ja-jp/office/vba/api/excel.xlautofilltype
on error resume next
Call App.Selection.AutoFill( Workbook.Sheets(1).Range("B2:B20"), 2 )
if Err.Number <> 0 then
	MsgBox( "ERROR : " & Err.Description )
	App.Quit()
	Wscript.Quit()
end if
on error goto 0

' ****************************
' 参照
' 最後の 1 は、使用するフィルター
' の番号です
' ****************************
FilePath = App.GetSaveAsFilename(,"Excel ファイル (*.xlsx), *.xlsx", 1)
if FilePath = "False" Then
	MsgBox "Excel ファイルの保存選択がキャンセルされました"
	WorkBook.Saved = True
	App.Quit()
	Wscript.Quit()
End If

' ****************************
' 保存
' 拡張子を .xls で保存するには
' Call ExcelBook.SaveAs( BookPath, 56 ) とします
' ****************************
on error resume next
Workbook.SaveAs( FilePath )
if Err.Number <> 0 then
	MsgBox( "ERROR : " & Err.Description )
	App.Quit()
	Wscript.Quit()
end if
on error goto 0

' ****************************
' 終了
' ****************************
App.Quit()

MsgBox( "処理が終了しました" )
