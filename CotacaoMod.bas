Attribute VB_Name = "CotacaoMod"
Option Explicit
Function Cotacao(ByVal Code As String, Optional Prop As String = "Ultimo")
    'Application.Volatile (True)
    
    Dim xmlhttp
    Set xmlhttp = CreateObject("msxml2.xmlhttp")
    xmlhttp.Open "POST", "http://www.bmfbovespa.com.br/Pregao-Online/ExecutaAcaoAjax.asp?CodigoPapel=" & Code, False
    xmlhttp.setrequestheader "Content-Type", "application/x-www-form-urlencoded"
    xmlhttp.send ""
    Dim result
    Set result = xmlhttp.responseXml.getElementsByTagName("Papel").NextNode
    Cotacao = result.getAttribute(Prop)
End Function

Public Sub RefreshCotacao()
    Dim cell
    For Each cell In Selection
        cell.Formula = Replace(cell.Value, "Cotacao", "Cotacao", 1, 1, vbTextCompare)
    Next cell
End Sub

