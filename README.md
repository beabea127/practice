# practice
practice

Option Explicit

Sub TestDownloadPDF_C3toC5()

    Dim ws As Worksheet
    Dim row As Long
    Dim url As String
    Dim tmpFolder As String
    Dim savePath As String

    Set ws = ActiveSheet

    ' 一時保存先フォルダ（%TEMP%\PdfTest\）
    tmpFolder = Environ$("TEMP") & "\PdfTest\"
    If Dir(tmpFolder, vbDirectory) = "" Then
        MkDir tmpFolder
    End If

    ' C3〜C5 のループ
    For row = 3 To 5
        
        url = Trim(CStr(ws.Cells(row, "C").Value))
        
        If url <> "" Then
            savePath = tmpFolder & "pdf_test_" & row & ".pdf"
            
            If DownloadFileSimple(url, savePath) Then
                MsgBox "ダウンロード成功: " & savePath, vbInformation
            Else
                MsgBox "ダウンロード失敗: " & url, vbExclamation
            End If
        End If

    Next row

End Sub


' URL→ファイル保存の簡易版関数（PDFダウンロード用）
Private Function DownloadFileSimple(ByVal url As String, ByVal savePath As String) As Boolean
    On Error GoTo ErrHandler

    Dim http As Object
    Dim stream As Object

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.send

    If http.Status <> 200 Then
        DownloadFileSimple = False
        Exit Function
    End If

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile savePath, 2
    stream.Close

    DownloadFileSimple = True
    Exit Function

ErrHandler:
    DownloadFileSimple = False
End Function
