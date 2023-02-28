Attribute VB_Name = "Scraper"
Option Explicit
Dim TotalPagesScraped As Integer

Public Function scrapper()

Dim ctimer As New ctimer
ctimer.StartCounter

Dim http As New MSXML2.XMLHTTP60
Dim doc As New HTMLDocument
Dim ecoll As Object
Dim item As Object
Dim URL As String
Dim page As Integer
Dim tbl As ListObject

Set tbl = Range("ScrapedData").ListObject
If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete

TotalPagesScraped = 0

For page = 1 To 10
    Debug.Print "Scrapping page " & page
    
    http.Open "GET", "https://www.internsg.com/jobs/" & page, True
    http.send
    
    Do While http.ReadyState <> 4
        DoEvents
    Loop

    doc.body.innerHTML = http.responseText
    Set ecoll = doc.getElementsByClassName("ast-row")
    
    For Each item In ecoll
        URL = ExtractURL(item.innerHTML)
        If Len(URL) > 0 Then
            ScrapPage URL
        End If
    Next item
    
Next page

Debug.Print "Total time taken: " & ctimer.TimeElapsed

End Function
Private Sub ScrapPage(ByVal URL As String)

Dim http As New MSXML2.XMLHTTP60
Dim doc As New HTMLDocument
Dim ecoll As Object
Dim item As Object
Dim Details As Variant
Dim HeaderIndex As New Collection
Dim tmpArr(1 To 14) As String
Dim tbl As ListObject
Dim row As ListRow

With HeaderIndex
    .Add 1, "Company"
    .Add 2, "Designation"
    .Add 3, "Date Listed"
    .Add 4, "Job Type"
    .Add 5, "Job Period"
    .Add 6, "Profession"
    .Add 7, "Industry"
    .Add 8, "Location Name"
    .Add 9, "Address"
    .Add 10, "Allowance / Remuneration"
    .Add 11, "Company Profile"
    .Add 12, "Job Description"
    .Add 13, "Application Instructions"
End With

Set tbl = Range("ScrapedData").ListObject

http.Open "GET", URL, True
http.send

Do While http.ReadyState <> 4
    DoEvents
Loop

doc.body.innerHTML = http.responseText
Set ecoll = doc.getElementsByClassName("ast-row p-3")

For Each item In ecoll
    Details = ExtractDetails(item.innerHTML)
    If Len(Details(0)) > 0 And Len(Details(1)) > 0 Then
        tmpArr(HeaderIndex(Details(0))) = Details(1)
    End If
Next item

tmpArr(14) = URL

Set row = tbl.ListRows.Add
row.Range = tmpArr
TotalPagesScraped = TotalPagesScraped + 1

Debug.Print "Total pages scraped: " & TotalPagesScraped

End Sub
Private Function ExtractURL(ByVal strInput As String) As String

On Error Resume Next

Dim regex As New RegExp
Dim strPattern As String
Dim match As Variant

If IsNull(strInput) Or strInput = "" Then
    Exit Function
End If

strPattern = "<div class=""ast-col-lg-1""><a href=""(.*?)""><span class=""text-monospace"

With regex
    .Global = True
    .MultiLine = True
    .IgnoreCase = False
    .Pattern = strPattern
End With

Set match = regex.Execute(strInput)
ExtractURL = match(0).SubMatches(0)

End Function
Private Function ExtractDetails(ByVal strInput As String) As Variant

On Error Resume Next

Dim regex As New RegExp
Dim match As Variant
Dim strPatternHeader As String
Dim strPatternContent As String
Dim strHeader As String
Dim strContent As String

If IsNull(strInput) Or strInput = "" Then
    Exit Function
End If

strInput = Replace(Replace(strInput, vbLf, ""), "  ", "")

strPatternHeader = "<div class=""ast-col-md-2 font-weight-bold"">(.*?)<\/?div>"

With regex
    .Global = True
    .MultiLine = True
    .IgnoreCase = False
    
    .Pattern = strPatternHeader
    Set match = .Execute(strInput)
    strHeader = match(0).SubMatches(0)
    
    If strHeader = "Profession" Or strHeader = "Industry" Then
        strPatternContent = "<i class=""gg-bell""><\/i><\/span>(.*?)<\/?div>"
    Else
        strPatternContent = "<div class=""ast-col-md-10"">(.*?)<\/?div>"
    End If
    
    .Pattern = strPatternContent
    Set match = .Execute(strInput)
    strContent = match(0).SubMatches(0)
    
    .Pattern = "<p>"
    strContent = .Replace(strContent, vbNewLine)
    
    .Pattern = "<li>"
    strContent = .Replace(strContent, vbNewLine & Chr(149))
    
    .Pattern = "<br>"
    strContent = .Replace(strContent, vbNewLine & vbNewLine)
    
    .Pattern = "<.*?>"
    strContent = .Replace(strContent, "")
    
    .Pattern = "&nbsp;"
    strContent = .Replace(strContent, "")

End With

ExtractDetails = Array(strHeader, strContent)

End Function
