' ****************************
' Excel �I�u�W�F�N�g�쐬
' ****************************
Set App = CreateObject("Excel.Application")

' ****************************
' �x�����o���Ȃ��悤�ɂ���
' ****************************
App.DisplayAlerts = False

' ****************************
' �u�b�N�ǉ�
' ****************************
App.Workbooks.Add()

' ****************************
' �ǉ������u�b�N���擾
' ****************************
Set Workbook = App.Workbooks( App.Workbooks.Count )

' ****************************
' ����A�u�b�N�ɂ̓V�[�g���
' �Ƃ����O��ŏ������Ă��܂���
' �K�v�ł���΁ABook.Worksheets.Count
' �Ō��݂̃V�[�g�̐����擾�ł��܂�
' ****************************
Set Worksheet = Workbook.Worksheets( 1 )
Worksheet.Activate()

' ****************************
' Add �ł� �������Ɏw�肵��
' �I�u�W�F�N�g�̃V�[�g�̒���ɁA
' �V�����V�[�g��ǉ����܂��B
' ****************************
Call Workbook.Worksheets.Add(,Worksheet)

' ****************************
' �V�[�g���ݒ�
' ****************************
Workbook.Sheets(1).Name = "�����V�[�g"
Workbook.Sheets(2).Name = "�ǉ��V�[�g"

' ****************************
' �f�[�^����
' ****************************
Workbook.Sheets(1).Activate()
Workbook.Sheets(1).Cells(1, 2) = "�Ј��R�[�h"
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
' �Q��
' �Ō�� 1 �́A�g�p����t�B���^�[
' �̔ԍ��ł�
' ****************************
FilePath = App.GetSaveAsFilename(,"Excel �t�@�C�� (*.xlsx), *.xlsx", 1)
if FilePath = "False" Then
	MsgBox "Excel �t�@�C���̕ۑ��I�����L�����Z������܂���"
	WorkBook.Saved = True
	App.Quit()
	Wscript.Quit()
End If

' ****************************
' �ۑ�
' �g���q�� .xls �ŕۑ�����ɂ�
' Call ExcelBook.SaveAs( BookPath, 56 ) �Ƃ��܂�
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
' �I��
' ****************************
App.Quit()

MsgBox( "�������I�����܂���" )
