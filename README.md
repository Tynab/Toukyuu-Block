# TOUKYUU BLOCK TOOL
Tool to help 西山 team of エマール group calculate ブロック of 東急 from 文化シャッター partner.

## MASK
<p align="center">
<img src="https://raw.githubusercontent.com/Tynab/Toukyuu-Block/main/pic/0.png"></img>
</p>

## CODE DEMO
```vb
''' <summary>
''' Check internet connection.
''' </summary>
''' <returns>Connection state.</returns>
Private Function IsNetAvail()
    Dim objResp As WebResponse
    Try
        objResp = WebRequest.Create(New Uri(My.Resources.link_base)).GetResponse
        objResp.Close()
        objResp = Nothing
        Return True
    Catch ex As Exception
        Return False
    End Try
End Function
```