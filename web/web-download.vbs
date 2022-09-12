set objSrvHTTP = Wscript.CreateObject("Msxml2.ServerXMLHTTP")
set Stream = Wscript.CreateObject("ADODB.Stream")

dim url : url = "https://github.com/winofsql/subject2/archive/refs/heads/main.zip"

call objSrvHTTP.Open("GET", url, False )
objSrvHTTP.Send
Stream.Open
Stream.Type = 1
Stream.Write objSrvHTTP.responseBody
Stream.SaveToFile "subject2-main.zip", 2
Stream.Close
