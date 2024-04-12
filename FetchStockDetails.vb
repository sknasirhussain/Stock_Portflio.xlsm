Option Explicit

Sub GetStockDetails()

    Dim stockSymbolRange As Range
    Dim cell As Range
    Dim quoteApiUrl As String
    Dim profileApiUrl As String
    Dim apiKey As String
    Dim httpRequest As Object
    Dim response As String
    Dim jsonResponse As Object
    Dim currentPrice As Double
    Dim High As Double
    Dim Low As Double
    Dim Name As String
    
    quoteApiUrl = "https://finnhub.io/api/v1/quote"
    profileApiUrl = "https://finnhub.io/api/v1/stock/profile2"
    apiKey = "cnsiq61r01qtn496p0o0cnsiq61r01qtn496p0og"
    
    Set stockSymbolRange = Range("B3:B22")
    
    For Each cell In stockSymbolRange
  
        If cell.Value <> "" Then
        
            quoteApiUrl = "https://finnhub.io/api/v1/quote?symbol=" & cell.Value & "&token=" & apiKey
            profileApiUrl = "https://finnhub.io/api/v1/stock/profile2?symbol=" & cell.Value & "&token=" & apiKey
            
            Set httpRequest = CreateObject("MSXML2.XMLHTTP")
            httpRequest.Open "GET", quoteApiUrl, False
            httpRequest.send
            
            If httpRequest.Status = 200 Then
                response = httpRequest.responseText
                Set jsonResponse = JsonConverter.ParseJson(response)
                currentPrice = jsonResponse("c")
                High = jsonResponse("h")
                Low = jsonResponse("l")
            Else
                cell.Offset(0, 2).Value = "Error: " & httpRequest.Status & " - " & httpRequest.statusText
                GoTo CleanUp
            End If
   
            Set httpRequest = CreateObject("MSXML2.XMLHTTP")
            httpRequest.Open "GET", profileApiUrl, False
            httpRequest.send
            
            If httpRequest.Status = 200 Then
                response = httpRequest.responseText
                Set jsonResponse = JsonConverter.ParseJson(response)
                Name = jsonResponse("name")
            Else
                cell.Offset(0, 1).Value = "Error: " & httpRequest.Status & " - " & httpRequest.statusText
                GoTo CleanUp
            End If
      
            cell.Offset(0, 1).Value = Name
            cell.Offset(0, 2).Value = currentPrice
            cell.Offset(0, 8).Value = High
            cell.Offset(0, 9).Value = Low
            
CleanUp:
            Set httpRequest = Nothing
        End If

        quoteApiUrl = "https://finnhub.io/api/v1/quote"
        profileApiUrl = "https://finnhub.io/api/v1/stock/profile2"
    Next cell

End Sub
