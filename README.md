# vbscript-startup

- ### [VBScript : 関数](https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc392480(v=msdn.10))
  - ### [MsgBox](https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc410277(v=msdn.10))



### index.hta
```
<html>

<head>
    <meta http-equiv="x-ua-compatible" content="ie=10">
    <meta http-equiv="content-type" content="text/html; charset=UTF-8">
    <script language="VBScript">
Function Test()

        Dim message: message = document.getElementById("item1").value
        MsgBox(message)

End Function
    </script>
</head>

<body>

    <input type="text" id="item1">
    <input type="button" id="btn1" value="実行" onclick='Call Test()'>

</body>

</html>
```
