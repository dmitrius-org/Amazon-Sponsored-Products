Attribute VB_Name = "mdlAmazonSPR"
Const server_name = ""
Const database_name = ""
Const User_ID = ""
Const Password = ""

Sub AmazonInbox(ByVal EntryIDCollection As String)
' Sponsored Products Beworbenes Produkt Bericht
' ������������ �������� ����� � ������������ ���������
' ��������� ��������� �� ������ ���� excel � ��������� ������ � ���� ������

Dim bodyString As String
Dim bodyStringLines
Dim splitLine
Dim hyperlink As String
Dim myItem As MailItem
'

' Dim olApp As Outlook.Application
' Dim objNS As Outlook.NameSpace
'Set olApp = Outlook.Application
'Set objNS = olApp.GetNamespace("MAPI")
'Set Items = objNS.GetDefaultFolder(olFolderInbox).Items
 
Set myItem = Outlook.GetNamespace("MAPI").GetItemFromID(EntryIDCollection)

If myItem Is Nothing Then Exit Sub
'If myItems.Sent Then Exit Sub  'ignore sent items
 
Debug.Print myItem.Sent
Debug.Print myItem.SentOn
Debug.Print "Begin AmazonInbox"

If InStr(1, myItem.Body, "Sponsored Products Beworbenes Produkt Bericht", vbTextCompare) > 0 Then
    
    If InStr(1, myItem.Body, "Download", vbTextCompare) > 0 Then
        bodyString = myItem.Body
        ' ��������� ������ �� �������
        bodyStringLines = Split(bodyString, vbCrLf)
        
        For Each splitLine In bodyStringLines ' ���� �� �������
            keyStart = InStr(splitLine, "Download")
                If keyStart > 0 Then
                    On Error Resume Next
                    '�������� ������ �� ��� ������ Download
                    hyperlink = Split(Split(splitLine, "<")(1), ">")(0)
                    If Len(hyperlink) > 0 Then
                      Exit For
                    End If
                    On Error GoTo 0
                End If
        Next
        
        Debug.Print "hyperlink is", hyperlink
       
        If Len(hyperlink) = 0 Then Exit Sub

        Dim fFile As String
        fFile = DownloadFile(hyperlink)
        If fFile <> "" Then
          ImportDB fFile
        End If
        
    End If
    
End If

End Sub

Function DownloadFile(myurl As String) As String
' DownloadFile ���������� ����� �� ������
' myurl - ������ �� ����
'
  DownloadFile = ""
  Dim saveDirectoryPath, fileName As String
  fileName = "tmp_amazon.xlsx"

  Set fs = CreateObject("Scripting.FileSystemObject")
  Set Folder = fs.GetSpecialFolder(2) ' 2 - TemporaryFolder
  saveDirectoryPath = Folder.Path

  Dim WinHttpReq As Object
  Set WinHttpReq = CreateObject("Msxml2.ServerXMLHTTP.6.0")
    
  WinHttpReq.Open "GET", myurl, False
  WinHttpReq.setRequestHeader "Content-Type", "text/json"
  WinHttpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  WinHttpReq.setRequestHeader "User-Agent", "Mozilla/5.0 (iPad; U; CPU OS 3_2_1 like Mac OS X; en-us) AppleWebKit/531.21.10 (KHTML, like Gecko) Mobile/7B405"
  WinHttpReq.Send
  
  Debug.Print "Download status", WinHttpReq.Status
  
  If WinHttpReq.Status = 200 Then
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1
    oStream.Write WinHttpReq.responseBody
    oStream.SaveToFile saveDirectoryPath & "\" & fileName, 2 ' 1 = no overwrite, 2 = overwrite
    oStream.Close
    
    DownloadFile = saveDirectoryPath & "\" & fileName
    Debug.Print "DownloadFile", DownloadFile
  End If
End Function

Function ImportDB(aFileName As String) As Boolean '
Dim conDB As ADODB.connection
Dim conXls As ADODB.connection
Dim rstDB As ADODB.Recordset
Dim rsXls As ADODB.Recordset

  On Error GoTo err
  
  Set conDB = New ADODB.connection
   
  ImportDB = False
  
  If aFileName = "" Then
    Debug.Print "������, �� ����� �������� ������ ��� ��������!"
    Exit Function
  End If
  
  ' ����������� � ��
  With conDB
    'SQLOLEDB.1
    '.ConnectionString = "Provider=SQLOLEDB" & _
      ";Data Source=" & server_name & _
      ";Initial Catalog=" & database_name & _
      ";ID=" & User_ID & _
      ";PWD=" & Password & _
      ";" 'Integrated Security=SSPI<�>
     'SQLNCLI11
     .ConnectionString = "Provider=SQLNCLI11" & _
      ";Server=" & server_name & _
      ";database=" & database_name & _
      ";Uid=" & User_ID & _
      ";Pwd=" & Password & _
      ";"
    .ConnectionTimeout = 10
    .Open
  End With
  'On Error GoTo 0
  
  If conDB.State = 1 Then
    Debug.Print "Connected to DB!"
    'Application.StatusBar = "Connected!"
  Else
    GoTo Exit_
  End If
  
  'On Error GoTo Exit_
  '���������, ��� �� ���� ��� ��������, ���� ������ �� ���������, �� ��� ����������
  If aFileName <> "" Then
    '������� ����������� � ���������
    Set conXls = New ADODB.connection 'CreateObject("ADODB.Connection")
    conXls.Provider = "Microsoft.ACE.OLEDB.12.0"
    conXls.Properties("Data Source") = aFileName
    '������ ��� �����
    conXls.Properties("Extended Properties") = "Excel 12.0 Xml; HDR=YES; IMEX=1;"
    '������ �������
    strSQL = "select * from [Sponsored Product Advertised Pr$]"
    '������������� ����������� � ���������
    conXls.Open
    '������� Recordset ��� ������ �� ���������
    Set rsXls = New ADODB.Recordset
    '��������� � ��������� ������ �� ���������
    rsXls.Open strSQL, conXls
   
    If rsXls.State = 1 Then
      Debug.Print "Connected to EXCEL!"
    Else
      GoTo Exit_
    End If
  
    '������� Recordset ��� ������ �� ����
    Set rstDB = New ADODB.Recordset
    
    ' �������� �� ����������� �� ������ �����
    rstDB.Open "Select * From dbo.importTest (nolock) where Datum = convert(date, '" & rsXls.Fields(0).Value & "', 104)", conDB, adOpenKeyset, adLockOptimistic
    If rstDB.RecordCount > 0 Then
      Debug.Print "������ ��� �����������!"
      GoTo Exit_
    End If
    
    rstDB.Close
    '��������� ������ Recordset
    rstDB.Open "dbo.importTest", conDB, adOpenDynamic, adLockOptimistic, adCmdTable
    '���������� ��� �������� ���������� ������������� �������
    counter = 0
    
    '��������� ���� ��� �������� ������ �� ��������� � ����,
    '�.�. ���� �������� ����� ��� ���� ������ � ����� � ���� ������
    While Not (rsXls.EOF)
      '����������� �������� �� ������� ��������� ������� � ���� ������
     ' Debug.Print rsXls.Fields(9).Type, rsXls.Fields(17).Value, rsXls.Fields(17).Type, IsNull(rsXls.Fields(17).Value)

      With rstDB
        .AddNew
        .Fields("Datum") = rsXls.Fields(0).Value 'rsXls.Fields("Datum").Value
        .Fields("Portfolioname") = rsXls.Fields(1).Value
        .Fields("Wahrung") = rsXls.Fields(2).Value
        .Fields("KampagnenName") = rsXls.Fields(3).Value
        .Fields("Anzeigengruppenname") = rsXls.Fields(4).Value
        .Fields("SKU") = rsXls.Fields(5).Value
        .Fields("ASIN") = rsXls.Fields(6).Value
        .Fields("Impressionen") = rsXls.Fields(7).Value
        .Fields("Klicks") = rsXls.Fields(8).Value
        .Fields("Klickrate") = CSPrcToDbl(rsXls.Fields(9).Value)
        .Fields("KlickCPC") = rsXls.Fields(10).Value
        .Fields("Ausgaben") = rsXls.Fields(11).Value
        .Fields("UmsatzGesamt") = rsXls.Fields(12).Value
        .Fields("ACOS") = CSPrcToDbl(rsXls.Fields(13).Value)
        .Fields("ROAS") = rsXls.Fields(14).Value
        .Fields("AuftrageGesamt") = rsXls.Fields(15).Value
        .Fields("EinheitenGesamt") = rsXls.Fields(16).Value
        .Fields("Konversionsrate").Value = CSPrcToDbl(rsXls.Fields(17).Value)
        .Fields("BeworbeneSKUEinheiten") = rsXls.Fields(18).Value
        .Fields("AndereSKUEinheiten") = rsXls.Fields(19).Value
        .Fields("BeworbeneSKUUmsatze") = rsXls.Fields(20).Value
        .Fields("AndereSKUUmsatze") = rsXls.Fields(21).Value
        .Update
        '����������� ��� �������
       counter = counter + 1
      End With
      '������ ��������� ������
      rsXls.MoveNext
    Wend
  
    '��������� ����������� � ���� MSSql
    conDB.Close
    Set conDB = Nothing
    '��������� �������� ������
    conXls.Close
    Set conXls = Nothing
    
    '������� Recordset
    Set rstDB = Nothing
    Set rsXls = Nothing
    '� ������� �� �����, ������� �� ������������� �����
    'MsgBox counter
    Debug.Print "Imported", aFileName
    Debug.Print "Imported rows", counter
  End If
Exit_:
Exit Function
err:
  Debug.Print "��������� ������: " & err.Description
End Function

Function CSPrcToDbl(val As Variant) As Variant
 
 

 If IsNull(val) Then
   CSPrcToDbl = val
 Else
    Select Case VarType(val)
        Case 5
            CSPrcToDbl = val * 100#
            ' The following is the only Case clause that evaluates to True.
        Case 8
            CSPrcToDbl = Replace(val, "%", "")
        Case Else
            CSPrcToDbl = val
    End Select
   
 End If
 
 'Debug.Print VarType(val), val, CSPrcToDbl
End Function

Sub testImportDB()
  ImportDB "C:\Users\dmital\AppData\Local\Temp\tmp_amazon.xlsx"
End Sub

Sub Search_Inbox() ' �� ������������
' Sponsored Products Beworbenes Produkt Bericht
' ������������ �������� ����� � ������������ ���������

Dim olFolder As Outlook.Folder
Dim myItems As Outlook.Items
Dim bodyString As String
Dim bodyStringLines
Dim splitLine
Dim hyperlink As String

Set olFolder = Application.GetNamespace("MAPI").Folders("dmital@rdb.ru").Folders("��������").Folders("test")

Set myItems = olFolder.Items
'
For Each myItem In myItems
    If InStr(1, myItem.Body, "Amazon Ads", vbTextCompare) > 0 Then
        If InStr(1, myItem.Body, "Download", vbTextCompare) > 0 Then
        bodyString = myItem.Body
        ' ��������� ������ �� �������
        bodyStringLines = Split(bodyString, vbCrLf)
        
        For Each splitLine In bodyStringLines ' ���� �� �������
            keyStart = InStr(splitLine, "Download")
                If keyStart > 0 Then
                    On Error Resume Next
                    '�������� ������ �� ��� ������ Download
                    hyperlink = Split(Split(splitLine, "<")(1), ">")(0)
                    If Len(hyperlink) > 0 Then
                      Exit For
                    End If
                    On Error GoTo 0
                End If
        Next
        Debug.Print "hyperlink is", hyperlink
       
        If Len(hyperlink) = 0 Then Exit Sub
        
        Dim fFile As String
        'Dim fResult As Boolean
        fFile = DownloadFile(hyperlink)
        If fFile <> "" Then
          'fResult =
          ImportDB fFile
        End If
        
        'DownloadFile1 (hyperlink)
        End If
    Else
        found = False
    End If
Next

Set olFolder = Nothing
Set myItems = Nothing

End Sub

