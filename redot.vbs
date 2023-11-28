' カレントディレクトリの取得
Dim currentDirectory

' 対象のファイル一覧を取得
Dim fso, folder, files, file
Set fso = CreateObject("Scripting.FileSystemObject")
currentDirectory = fso.GetAbsolutePathName(".")
Set folder = fso.GetFolder(currentDirectory)
Set files = folder.Files

' 対象ファイルの処理
For Each file In files
    ' 拡張子が .png のファイルを処理
    If LCase(fso.GetExtensionName(file.Path)) = "png" Then
        ' ファイル名から .(ドット) を削除
        Dim fileName, newFileName
        fileName = fso.GetFileName(file.Path)
        fileName = LCase(fileName)
        newFileName = Replace(fileName, ".png", "")
        newFileName = Replace(newFileName, ".", "")
        newFileName = newFileName & ".png"
        
        ' ファイル名変更
        If fileName <> newFileName Then
            Dim newPath
            newPath = fso.BuildPath(currentDirectory, newFileName)
            fso.MoveFile file.Path, newPath
        End If
    End If
Next
