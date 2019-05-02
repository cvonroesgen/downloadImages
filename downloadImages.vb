Imports System.Xml
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Module downloadImages
    Dim NameValues As New Dictionary(Of String, String)
    Sub Main()
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        Dim i As Integer
        Dim arguments As String() = Environment.GetCommandLineArgs()
        If arguments.Length < 2 Then
            Console.WriteLine("Please include one command line argument with the following values:")
            Console.WriteLine("")
            Console.WriteLine("dbid: the table identifier of a Quick Base table")
            Console.WriteLine("qid: the report identifier of a report in the above Quick Base table")
            Console.WriteLine("filenameTemplate: the template of how to name a file attachment downloaded from QuickBase")
            Console.WriteLine("usertoken: the Quick Base usertoken used to authenticate")
            Console.WriteLine("server: the Quick Base server i.e. kgiwireless.quickbase.com")
            Console.WriteLine("folder: the location to place the downloaded pictures")
            Console.WriteLine("")
            Console.WriteLine("for example:")
            Console.WriteLine("downloadImages ""dbid=bpedrarn2&qid=11&filenameTemplate=42_0_52&usertoken=lsakdbvcayreliawucbrlwuie&server=kgiwireless.quickbase.com&folder=W:\\folder1\folder2\folder3""")
            Exit Sub
        Else
            Dim NameValuesPairs() As String = arguments(1).Split("&")
            For i = 0 To NameValuesPairs.Length - 1
                Dim nameValue() As String = NameValuesPairs(i).Split("=")
                NameValues.Add(nameValue(0), nameValue(1))
            Next
        End If

        Dim filenameTemplate As String = NameValues("filenameTemplate")
        Dim filenameTemplateFids() As String = filenameTemplate.Split("_")

        Dim qdb As New QuickBaseClient(NameValues("usertoken"))
        qdb.setServer(NameValues("server"), True)
        Dim recordCount As Integer = 0
        Dim skip As Integer = 0
        Dim chunkSize = 100
        Do
            Dim xmlRecords As XmlDocument = qdb.DoQuery(NameValues("dbid"), NameValues("qid"), "a", "3", "skp-" & skip & ".num-" & chunkSize)
            Dim fields As XmlNodeList = xmlRecords.SelectNodes("/*/table/fields/field[@field_type='file']")
            Dim getFileExtension As Regex = New Regex("(\.[0-9a-z]+)$", RegexOptions.IgnoreCase)

            Dim records As XmlNodeList = xmlRecords.SelectNodes("/*/table/records/record")
            recordCount = records.Count
            Dim j As Integer
            For i = 0 To records.Count - 1
                Dim rid As String = records(i).SelectSingleNode("@rid").InnerText()
                For j = 0 To fields.Count - 1
                    Dim delimiter = ""
                    Dim filename As String = ""
                    Dim k As Integer
                    For k = 0 To filenameTemplateFids.Count - 1
                        filename &= delimiter
                        delimiter = "_"
                        If filenameTemplateFids(k) = "0" Then
                            filename &= fields(j).SelectSingleNode("label").InnerText()
                        Else
                            filename &= records(i).SelectSingleNode("f[@id=" & filenameTemplateFids(k) & "]").InnerText()
                        End If
                    Next
                    'now we need to download the file
                    Dim fid As String = fields(j).SelectSingleNode("@id").InnerText()
                    Dim urlNode As XmlNode = records(i).SelectSingleNode("f[@id='" & fid & "']/url")
                    If urlNode IsNot Nothing Then
                        Dim url As String = urlNode.InnerText()
                        Dim extensionMatch As Match = getFileExtension.Match(url)
                        filename &= extensionMatch.Value
                        Console.WriteLine("Downloading: " & filename)
                        qdb.downloadAttachedFile(NameValues("dbid"), rid, fid, NameValues("folder"), filename)
                    End If
                Next
            Next
            skip += chunkSize
        Loop While recordCount > 0

    End Sub
    Public Class QuickBaseClient

        Private userToken As String
        Private strProxyPassword As String
        Private strProxyUsername As String
        Private ticket As String
        Private apptoken As String
        Private QDBHost As String = ""
        Private useHTTPS As Boolean = True
        Public GMTOffset As Single


        Public errorcode As Integer
        Public errortext As String
        Public errordetail As String
        Public httpContentLengthProgress As Integer
        Public httpContentLength As Integer

        Private Const OB32CHARACTERS As String = "abcdefghijkmnpqrstuvwxyz23456789"
        Private Const Map64 As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
        Private Const MILLISECONDS_IN_A_DAY As Double = 86400000.0#
        Private Const DAYS_BETWEEN_JAVASCRIPT_AND_MICROSOFT_DATE_REFERENCES As Double = 25569.0#
        Function makeCSVCells(ByRef cells As ArrayList) As String
            Dim i As Integer
            Dim cell As String
            makeCSVCells = ""
            For i = 0 To cells.Count - 1
                If cells(i) Is Nothing Then
                    cell = ""
                Else
                    cell = cells(i).ToString()
                End If
                makeCSVCells = makeCSVCells & """" & cell.Replace("""", """""") & """, "
            Next
        End Function

        Function encode32(ByVal strDecimal As String) As String

            Dim ob32 As String = ""
            Dim intDecimal As Integer
            intDecimal = CInt(strDecimal)
            Dim remainder As Integer

            Do While (intDecimal > 0)
                remainder = intDecimal Mod 32
                ob32 = Mid(OB32CHARACTERS, CInt(remainder) + 1, 1) & ob32
                intDecimal = intDecimal \ 32
            Loop
            encode32 = ob32

        End Function
        Public Function getTextByFID(ByRef recordNode As XmlNode, ByRef fid As String) As String
            Dim cell As XmlNode = recordNode.SelectSingleNode("f[@id=" & fid & "]")
            If cell Is Nothing Then
                Err.Raise(vbObjectError + 5, "QuickBase.QuickBaseClient", "Could Not find fid " & fid)
            End If
            getTextByFID = cell.InnerText
        End Function
        Public Function makeClist(ByVal fids As Hashtable) As String
            Dim period As String = ""
            makeClist = ""
            For Each fid As DictionaryEntry In fids
                makeClist = makeClist & period & fid.Value
                period = "."
            Next
        End Function
        Public Function makeClist(ByRef fids As ArrayList) As String
            Dim period As String = ""
            makeClist = ""
            Dim i As Integer
            For i = 0 To fids.Count - 1
                makeClist = makeClist & period & fids(i)
                period = "."
            Next
        End Function
        Public Function makeClist(ByRef fids() As String) As String
            Dim period As String = ""
            makeClist = ""
            Dim fid As String
            For Each fid In fids
                makeClist = makeClist & period & fid
                period = "."
            Next
        End Function
        Public Function FieldAddChoices(ByVal dbid As String, ByVal fid As String, ByVal ParamArray NameValues() As Object) As Integer
            Dim xmlQDBRequest As XmlDocument
            Dim firstfield As Integer
            Dim lastfield As Integer
            Dim i As Integer

            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "fid", fid)
            lastfield = UBound(NameValues)
            firstfield = LBound(NameValues)
            For i = firstfield To lastfield
                addParameter(xmlQDBRequest, "choice", CStr(NameValues(i)))
            Next i

            xmlQDBRequest = APIXMLPost(dbid, "API_FieldAddChoices", xmlQDBRequest, useHTTPS)
            FieldAddChoices = CInt(xmlQDBRequest.DocumentElement.SelectSingleNode("/*/numadded").InnerText)
        End Function

        Public Function CreateDatabase(ByVal dbname As String, ByVal dbdesc As String) As String
            Dim xmlQDBRequest As XmlDocument

            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "dbname", dbname)
            addParameter(xmlQDBRequest, "dbdesc", dbdesc)
            CreateDatabase = ""
            CreateDatabase = APIXMLPost("main", "API_CreateDatabase", xmlQDBRequest, useHTTPS).DocumentElement.SelectSingleNode("/*/dbid").InnerText
        End Function
        Public Sub DeleteDatabase(ByVal dbid As String)
            Dim xmlQDBRequest As XmlDocument

            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "dbid", dbid)
            Call APIXMLPost(dbid, "API_DeleteDatabase", xmlQDBRequest, useHTTPS)
        End Sub
        Public Function AddField(ByVal dbid As String, ByVal label As String, ByVal fieldtype As String, ByVal Formula As Boolean) As String
            Dim xmlQDBRequest As XmlDocument

            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "label", label)
            addParameter(xmlQDBRequest, "type", fieldtype)
            If Formula Then
                addParameter(xmlQDBRequest, "mode", "virtual")
            End If
            AddField = ""
            AddField = APIXMLPost(dbid, "API_AddField", xmlQDBRequest, useHTTPS).DocumentElement.SelectSingleNode("/*/fid").InnerText
        End Function
        Public Sub SetFieldProperties(ByVal dbid As String, ByVal fid As String, ByVal ParamArray NameValues() As Object)
            Dim xmlQDBRequest As XmlDocument
            Dim lastfield As Integer
            Dim firstfield As Integer
            Dim i As Integer

            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "fid", fid)
            lastfield = UBound(NameValues)
            firstfield = LBound(NameValues)
            i = 0
            For i = firstfield To lastfield Step 2
                addParameter(xmlQDBRequest, CStr(NameValues(i)), CStr(NameValues(i + 1)))
            Next i
            Call APIXMLPost(dbid, "API_SetFieldProperties", xmlQDBRequest, useHTTPS)
        End Sub


        Public Function proxyAuthenticate(ByVal strUsername As String, ByVal strPassword As String) As Integer
            strProxyUsername = strUsername
            strProxyPassword = strPassword
            proxyAuthenticate = 0
        End Function
        Function downloadAttachedFile(ByVal dbid As String, ByVal rid As String, ByVal fid As String, ByVal DownloadDirectory As String, ByVal Filename As String) As String
            Filename = makeValidFilename(Filename)
            downloadAttachedFile = HTTPPost(QDBHost, True, "/up/" & dbid & "/a/r" & rid & "/e" & fid & "/?usertoken=" & userToken, "text/html", "", DownloadDirectory & "\" & Filename)
        End Function
        Public Function FindDBByName(ByVal dbname As String) As String
            Dim xmlQDBRequest As XmlDocument

            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "dbname", dbname)
            FindDBByName = ""
            FindDBByName = APIXMLPost("main", "API_FindDBByName", xmlQDBRequest, useHTTPS).DocumentElement.SelectSingleNode("/*/dbid").InnerText
        End Function
        Public Function CloneDatabase(ByVal sourcedbid As String, ByVal Name As String, ByVal Description As String) As String
            Dim xmlQDBRequest As XmlDocument
            Dim xmlNewDBID As XmlNode
            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "newdbname", Name)
            addParameter(xmlQDBRequest, "newdbdesc", Description)
            CloneDatabase = ""
            xmlQDBRequest = APIXMLPost(sourcedbid, "API_CloneDatabase", xmlQDBRequest, useHTTPS)
            If Not xmlQDBRequest.HasChildNodes Then
                Err.Raise(vbObjectError + 5, "QuickBase.QuickBaseClient", "Please login with an user account that has permission to create applications.")
            End If
            xmlNewDBID = xmlQDBRequest.DocumentElement.SelectSingleNode("/*/newdbid")
            If xmlNewDBID Is Nothing Then
                Err.Raise(vbObjectError + 5, "QuickBase.QuickBaseClient", "Please login with an user account that has permission to create applications in only one billing account.")
            Else
                CloneDatabase = xmlNewDBID.InnerText
            End If
        End Function
        Public Function ImportFromCSV(ByVal dbid As String, ByVal CSV As String, ByVal clist As String, ByRef rids() As Integer, ByVal skipfirst As Boolean) As Integer
            Dim xmlQDBRequest As XmlDocument
            Dim RidNodeList As XmlNodeList 'XmlNodeList

            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "clist", clist)
            addParameter(xmlQDBRequest, "msInUTC", "1")
            If skipfirst Then
                addParameter(xmlQDBRequest, "skipfirst", "1")
            End If
            addCDATAParameter(xmlQDBRequest, "records_csv", CSV)
            xmlQDBRequest = APIXMLPost(dbid, "API_ImportFromCSV", xmlQDBRequest, useHTTPS)
            RidNodeList = xmlQDBRequest.SelectNodes("/*/rids/rid")
            Dim ridListLength As Integer
            Dim i As Integer
            ridListLength = RidNodeList.Count
            If ridListLength > 0 Then
                ReDim rids(ridListLength - 1)
                For i = 0 To ridListLength - 1
                    rids(i) = CInt(RidNodeList(i).InnerText)
                Next i
            End If
            On Error Resume Next
            ImportFromCSV = CInt(xmlQDBRequest.DocumentElement.SelectSingleNode("/*/num_recs_added").InnerText)
            ImportFromCSV = CInt(xmlQDBRequest.DocumentElement.SelectSingleNode("/*/num_recs_updated").InnerText)
            xmlQDBRequest = Nothing
        End Function
        Public Function AddRecordByArray(ByVal dbid As String, ByRef update_id As String, ByRef NameValues(,) As Object) As String
            Dim xmlQDBRequest As XmlDocument
            Dim firstfield As Integer
            Dim lastfield As Integer
            Dim i As Integer
            AddRecordByArray = ""

            xmlQDBRequest = InitXMLRequest()
            lastfield = UBound(NameValues, 2)
            firstfield = LBound(NameValues, 2)

            For i = firstfield To lastfield
                If IsDBNull(NameValues(0, i)) Then
                    Err.Raise(vbObjectError + 2, "QuickBase.QuickBaseClient", "AddRecordByArray: Please do use null for field names or fids")
                    Exit Function
                End If
                If IsDBNull(NameValues(1, i)) Then
                    NameValues(1, i) = CObj("")
                End If

                If (IsNumeric(NameValues(0, i)) And Not IsDate(NameValues(0, i))) Then
                    addFieldParameter(xmlQDBRequest, "fid", CStr(NameValues(0, i)), NameValues(1, i))
                Else
                    addFieldParameter(xmlQDBRequest, "name", makeAlphaNumLowerCase(CStr(NameValues(0, i))), NameValues(1, i))
                End If
            Next i
            xmlQDBRequest = APIXMLPost(dbid, "API_AddRecord", xmlQDBRequest, useHTTPS)
            update_id = xmlQDBRequest.DocumentElement.SelectSingleNode("/*/update_id").InnerText
            AddRecordByArray = xmlQDBRequest.DocumentElement.SelectSingleNode("/*/rid").InnerText
        End Function

        Public Function EditRecordByArray(ByVal dbid As String, ByVal rid As String, ByRef update_id As String, ByRef NameValues(,) As Object) As String
            Dim xmlQDBRequest As XmlDocument
            Dim firstfield As Integer
            Dim lastfield As Integer
            Dim i As Integer
            EditRecordByArray = ""

            xmlQDBRequest = InitXMLRequest()
            lastfield = UBound(NameValues, 2)
            firstfield = LBound(NameValues, 2)



            For i = firstfield To lastfield
                If (IsNumeric(NameValues(0, i)) And Not IsDate(NameValues(0, i))) Then
                    addFieldParameter(xmlQDBRequest, "fid", CStr(NameValues(0, i)), NameValues(1, i))
                Else
                    addFieldParameter(xmlQDBRequest, "name", makeAlphaNumLowerCase(CStr(NameValues(0, i))), NameValues(1, i))
                End If
            Next i
            addParameter(xmlQDBRequest, "rid", rid)
            If update_id <> "" Then
                addParameter(xmlQDBRequest, "update_id", update_id)
            End If
            EditRecordByArray = APIXMLPost(dbid, "API_EditRecord", xmlQDBRequest, useHTTPS).DocumentElement.SelectSingleNode("/*/update_id").InnerText
        End Function
        Public Function AddRecord(ByVal dbid As String, ByRef update_id As String, ByVal ParamArray NameValues() As Object) As String
            Dim xmlQDBRequest As XmlDocument
            Dim firstfield As Integer
            Dim lastfield As Integer
            Dim i As Integer
            AddRecord = ""

            xmlQDBRequest = InitXMLRequest()
            lastfield = UBound(NameValues)
            firstfield = LBound(NameValues)
            If ((lastfield - firstfield + 1) Mod 2) <> 0 Then
                Err.Raise(vbObjectError + 3, "QuickBase.QuickBaseClient", "AddRecord: Please use an even number of arguements after the DBID")
                Exit Function
            End If


            For i = firstfield To lastfield Step 2
                If (IsNumeric(NameValues(i)) And Not IsDate(NameValues(i))) Then
                    addFieldParameter(xmlQDBRequest, "fid", CStr(NameValues(i)), NameValues(i + 1))
                Else
                    addFieldParameter(xmlQDBRequest, "name", makeAlphaNumLowerCase(CStr(NameValues(i))), NameValues(i + 1))
                End If
            Next i
            xmlQDBRequest = APIXMLPost(dbid, "API_AddRecord", xmlQDBRequest, useHTTPS)
            update_id = xmlQDBRequest.DocumentElement.SelectSingleNode("/*/update_id").InnerText
            AddRecord = xmlQDBRequest.DocumentElement.SelectSingleNode("/*/rid").InnerText
        End Function
        Public Function EditRecord(ByVal dbid As String, ByVal rid As String, ByRef update_id As String, ByVal ParamArray NameValues() As Object) As String
            Dim xmlQDBRequest As XmlDocument
            Dim firstfield As Integer
            Dim lastfield As Integer
            Dim i As Integer
            EditRecord = ""

            xmlQDBRequest = InitXMLRequest()
            lastfield = UBound(NameValues)
            firstfield = LBound(NameValues)
            If ((lastfield - firstfield + 1) Mod 2) <> 0 Then
                Err.Raise(vbObjectError + 4, "QuickBase.QuickBaseClient", "EditRecord: Please use an even number of arguements.")
                Exit Function
            End If


            For i = firstfield To lastfield Step 2
                If (IsNumeric(NameValues(i)) And Not IsDate(NameValues(i))) Then
                    addFieldParameter(xmlQDBRequest, "fid", CStr(NameValues(i)), NameValues(i + 1))
                Else
                    addFieldParameter(xmlQDBRequest, "name", makeAlphaNumLowerCase(CStr(NameValues(i))), NameValues(i + 1))
                End If
            Next i
            addParameter(xmlQDBRequest, "rid", rid)
            If update_id <> "" Then
                addParameter(xmlQDBRequest, "update_id", update_id)
            End If
            EditRecord = APIXMLPost(dbid, "API_EditRecord", xmlQDBRequest, useHTTPS).DocumentElement.SelectSingleNode("/*/update_id").InnerText

        End Function
        Public Function DeleteRecord(ByVal dbid As String, ByVal rid As Object) As String
            Dim xmlQDBRequest As XmlDocument
            DeleteRecord = ""

            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "rid", CStr(rid))

            DeleteRecord = APIXMLPost(dbid, "API_DeleteRecord", xmlQDBRequest, useHTTPS).DocumentElement.SelectSingleNode("/*/rid").InnerText

        End Function
        Public Function GetSchema(ByVal dbid As String) As XmlDocument
            Dim xmlQDBRequest As XmlDocument

            xmlQDBRequest = InitXMLRequest()
            GetSchema = APIXMLPost(dbid, "API_GetSchema", xmlQDBRequest, useHTTPS)
        End Function

        Public Function GetGrantedDBs(ByVal withEmbeddedTables As Boolean, ByVal excludeParents As Boolean, ByVal adminOnly As Boolean) As XmlDocument
            Dim xmlQDBRequest As XmlDocument

            xmlQDBRequest = InitXMLRequest()
            If withEmbeddedTables Then
                addParameter(xmlQDBRequest, "withEmbeddedTables", "1")
            Else
                addParameter(xmlQDBRequest, "withEmbeddedTables", "0")
            End If
            If excludeParents Then
                addParameter(xmlQDBRequest, "excludeParents", "1")
            Else
                addParameter(xmlQDBRequest, "excludeParents", "0")
            End If
            If adminOnly Then
                addParameter(xmlQDBRequest, "adminOnly", "1")
            End If
            addParameter(xmlQDBRequest, "realmAppsOnly", "true")
            GetGrantedDBs = APIXMLPost("main", "API_GrantedDBs", xmlQDBRequest, useHTTPS)
        End Function
        Public Function GetDBInfo(ByVal dbid As String) As XmlDocument
            Dim xmlQDBRequest As XmlDocument

            xmlQDBRequest = InitXMLRequest()
            GetDBInfo = APIXMLPost(dbid, "API_GetDBInfo", xmlQDBRequest, useHTTPS)
        End Function

        Public Function ChangeRecordOwner(ByVal dbid As String, ByVal rid As Object, ByVal Owner As String) As Boolean
            Dim xmlQDBRequest As XmlDocument

            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "rid", CStr(rid))
            addParameter(xmlQDBRequest, "newowner", Owner)
            On Error GoTo noChange
            Call APIXMLPost(dbid, "API_ChangeRecordOwner", xmlQDBRequest, useHTTPS)
            ChangeRecordOwner = True
            Exit Function

noChange:
            ChangeRecordOwner = False
            Exit Function
        End Function
        Public Function DoQuery(ByVal dbid As String, ByVal query As String, ByVal clist As String, ByVal slist As String, ByVal options As String) As XmlDocument
            Dim xmlQDBRequest As XmlDocument

            xmlQDBRequest = InitXMLRequest()
            If CStr(Val(query)) = query Then
                addParameter(xmlQDBRequest, "qid", query)
            Else
                addParameter(xmlQDBRequest, "query", query)
            End If
            addParameter(xmlQDBRequest, "clist", clist)
            addParameter(xmlQDBRequest, "slist", slist)
            addParameter(xmlQDBRequest, "options", options)
            addParameter(xmlQDBRequest, "fmt", "structured")
            addParameter(xmlQDBRequest, "includeRids", "1")
            DoQuery = APIXMLPost(dbid, "API_DoQuery", xmlQDBRequest, useHTTPS)
        End Function
        Public Function GenResultsTable(ByVal dbid As String, ByVal query As String, ByVal clist As String, ByVal slist As String, ByVal options As String) As String
            Dim xmlQDBRequest As XmlDocument

            xmlQDBRequest = InitXMLRequest()

            If Left(query, 1) = "{" And Right(query, 1) = "}" Then
                addParameter(xmlQDBRequest, "query", query)
            ElseIf CStr(Val(query)) = query Then
                addParameter(xmlQDBRequest, "qid", query)
            Else
                addParameter(xmlQDBRequest, "qname", query)
            End If
            addParameter(xmlQDBRequest, "clist", clist)
            addParameter(xmlQDBRequest, "slist", slist)
            addParameter(xmlQDBRequest, "options", options)
            GenResultsTable = APIHTMLPost(dbid, "API_GenResultsTable", xmlQDBRequest, useHTTPS)
        End Function

        Public Function DoQueryAsArray(ByVal dbid As String, ByVal query As String, ByVal clist As String, ByVal slist As String, ByVal options As String) As Object(,)
            Dim xmlQDBResponse As New XmlDocument
            xmlQDBResponse = DoQuery(dbid, query, clist, slist, options)
            Dim QDBRecord(,) As Object

            Dim i As Integer
            Dim j As Integer
            Dim FieldNodeList As XmlNodeList
            Dim RecordNodeList As XmlNodeList
            Dim intFields As Integer
            Dim strFieldValue As String

            FieldNodeList = xmlQDBResponse.DocumentElement.SelectNodes("/*/table/fields/field")
            intFields = FieldNodeList.Count

            RecordNodeList = xmlQDBResponse.DocumentElement.SelectNodes("/*/table/records/record")

            ReDim QDBRecord(RecordNodeList.Count + 1, intFields)

            For i = 0 To intFields - 1
                QDBRecord(0, i) = FieldNodeList(i).SelectSingleNode("label").InnerText
            Next i

            For i = 1 To RecordNodeList.Count
                For j = 0 To intFields - 1
                    On Error Resume Next
                    strFieldValue = RecordNodeList(i - 1).SelectSingleNode("f[" & CStr(j) & "]").InnerText
                    Select Case FieldNodeList(j).SelectSingleNode("@base_type").InnerText
                        Case "float"
                            If strFieldValue <> "" Then
                                QDBRecord(i, j) = makeDouble(strFieldValue)
                            End If
                        Case "text"
                            QDBRecord(i, j) = Replace(strFieldValue, Chr(10), vbCrLf)
                        Case "bool"
                            QDBRecord(i, j) = CBool(strFieldValue)
                        Case "int64"
                            If strFieldValue <> "" Then
                                If FieldNodeList(j).SelectSingleNode("@field_type").InnerText <> "date" Then
                                    QDBRecord(i, j) = CDbl(strFieldValue) / MILLISECONDS_IN_A_DAY
                                Else
                                    QDBRecord(i, j) = int64ToDate(strFieldValue)
                                End If
                            End If
                        Case "int32"
                            If FieldNodeList(j).SelectSingleNode("@field_type").InnerText = "userid" Then
                                On Error Resume Next
                                Dim tempLong As Integer
                                tempLong = CInt(strFieldValue)
                                If Err.Number = 0 Then
                                    QDBRecord(i, j) = xmlQDBResponse.SelectSingleNode("/*/table/lusers/luser[@id='" & strFieldValue & "']").InnerText
                                Else
                                    QDBRecord(i, j) = strFieldValue
                                End If
                                On Error GoTo 0
                            Else
                                QDBRecord(i, j) = CLng(strFieldValue)
                            End If
                    End Select
                Next j
            Next i
            DoQueryAsArray = QDBRecord
        End Function

        Public Function APIXMLPost(ByVal dbid As String, ByVal action As String, ByRef xmlQDBRequest As XmlDocument, ByVal useHTTPS As Boolean) As XmlDocument

            Dim script As String
            Dim content As String
            Dim req As HttpWebRequest
            Dim resp As HttpWebResponse
            Dim xmlStream As Stream
            Dim xmlTxtReader As XmlTextReader
            Dim xmlDoc As XmlDocument

            script = QDBHost & "/db/" & dbid & "?act=" & action
            If useHTTPS Then
                script = "https://" & script
            Else
                script = "http://" & script
            End If
            content = "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & xmlQDBRequest.OuterXml
            req = CType(WebRequest.Create(script), HttpWebRequest)
            req.ContentType = "text/xml"
            req.Method = "POST"
            Dim byteRequestArray As Byte() = Encoding.UTF8.GetBytes(content)
            req.ContentLength = byteRequestArray.Length
            Dim reqStream As Stream = req.GetRequestStream()
            reqStream.Write(byteRequestArray, 0, byteRequestArray.Length)
            resp = CType(req.GetResponse(), HttpWebResponse)
            reqStream.Close()
            'create a new stream that can be placed into an XmlTextReader
            xmlStream = resp.GetResponseStream()
            xmlTxtReader = New XmlTextReader(xmlStream)
            xmlTxtReader.XmlResolver = Nothing
            'create a new Xml document
            xmlDoc = New XmlDocument
            xmlDoc.Load(xmlTxtReader)
            xmlStream.Close()
            On Error Resume Next
            errorcode = CInt(resp.Headers("QUICKBASE-ERRCODE"))
            ticket = xmlDoc.DocumentElement.SelectSingleNode("/*/ticket").InnerText
            errortext = xmlDoc.DocumentElement.SelectSingleNode("/*/errtext").InnerText
            If xmlDoc.DocumentElement.SelectSingleNode("/*/errdetail") Is Nothing Then
                errordetail = xmlDoc.DocumentElement.SelectSingleNode("/*/errtext").InnerText
            Else
                errordetail = xmlDoc.DocumentElement.SelectSingleNode("/*/errdetail").InnerText
            End If
            On Error GoTo 0
            If errorcode <> 0 Then
                Err.Raise(vbObjectError + CInt(errorcode), "QuickBase.QuickBaseClient", script & ": " & errordetail)
            End If


            APIXMLPost = xmlDoc
        End Function
        Public Function APIHTMLPost(ByVal dbid As String, ByVal action As String, ByRef xmlQDBRequest As XmlDocument, ByVal useHTTPS As Boolean) As String

            Dim script As String


            script = "/db/" & dbid & "?act=" & action
            APIHTMLPost = HTTPPost(QDBHost, useHTTPS, script, "text/xml", xmlQDBRequest.OuterXml, "")

        End Function
        Private Function HTTPPost(ByVal QDBHost As String, ByVal useHTTPS As Boolean, ByVal script As String, ByVal contentType As String, ByVal content As String, ByVal fileName As String) As String
            Dim url As String
            Dim Client As WebClient

            Client = New WebClient
            Client.Headers.Add("Content-Type", contentType)
            url = QDBHost & script
            If useHTTPS Then
                url = "https://" & url
            Else
                url = "http://" & url
            End If
            Dim byteRequestArray As Byte() = Encoding.UTF8.GetBytes(content)

            Dim byteResponseArray As Byte() = Client.UploadData(url, "POST", byteRequestArray)
            If fileName = "" Then
                HTTPPost = Encoding.UTF8.GetString(byteResponseArray)
            Else
                'check if write file exists 
                If File.Exists(Path:=fileName) Then
                    'delete file
                    File.Delete(Path:=fileName)
                End If

                'create a fileStream instance to pass to BinaryWriter object
                Dim fsWrite As FileStream
                fsWrite = New FileStream(Path:=fileName,
                    mode:=FileMode.CreateNew, access:=FileAccess.Write)

                'create binary writer instance
                Dim bWrite As BinaryWriter
                bWrite = New BinaryWriter(output:=fsWrite)
                'write bytes out 
                bWrite.Write(byteResponseArray, 0, byteResponseArray.Length)


                'close the writer 
                bWrite.Close()

                fsWrite.Close()


                HTTPPost = fileName
            End If

        End Function
        Public Function setAppToken(ByVal aapptoken As String) As String
            apptoken = aapptoken
            setAppToken = apptoken
        End Function
        Public Function getServer() As String
            getServer = QDBHost
        End Function

        Public Function getTicket() As String
            If ticket = "" Then
                Dim xmlQDBRequest As XmlDocument

                xmlQDBRequest = InitXMLRequest()
                Call APIXMLPost("main", "API_Authenticate", xmlQDBRequest, useHTTPS)
            End If
            getTicket = ticket
        End Function

        Public Function InitXMLRequest() As XmlDocument
            Dim xmlQDBRequest As New XmlDocument
            Dim Root As XmlElement

            Root = xmlQDBRequest.CreateElement("qdbapi")
            xmlQDBRequest.AppendChild(Root)
            addParameter(xmlQDBRequest, "usertoken", userToken)
            InitXMLRequest = xmlQDBRequest
        End Function

        Public Sub addParameter(ByRef xmlQDBRequest As XmlDocument, ByVal Name As String, ByVal Value As String)
            Dim Root As XmlElement
            Dim ElementNode As XmlNode
            Dim TextNode As XmlNode

            Root = xmlQDBRequest.DocumentElement
            ElementNode = xmlQDBRequest.CreateNode(XmlNodeType.Element, Name, "")
            TextNode = xmlQDBRequest.CreateNode(XmlNodeType.Text, "", "")
            TextNode.InnerText = Value
            ElementNode.AppendChild(TextNode)
            Root.AppendChild(ElementNode)
            Root = Nothing
            ElementNode = Nothing
            TextNode = Nothing
        End Sub

        Public Sub addParameterWithAttribute(ByRef xmlQDBRequest As XmlDocument, ByVal Name As String, ByVal AttributeName As String, ByVal AttributeValue As String, ByVal Value As String)
            Dim Root As XmlElement
            Dim ElementNode As XmlNode
            Dim TextNode As XmlNode
            Dim Attribute As XmlAttribute

            Root = xmlQDBRequest.DocumentElement
            ElementNode = xmlQDBRequest.CreateNode(XmlNodeType.Element, Name, "")
            TextNode = xmlQDBRequest.CreateNode(XmlNodeType.Text, "", "")
            TextNode.InnerText = Value
            Attribute = xmlQDBRequest.CreateAttribute(AttributeName)
            Attribute.Value = AttributeValue
            ElementNode.Attributes.Append(Attribute)

            ElementNode.AppendChild(TextNode)
            Root.AppendChild(ElementNode)
            Root = Nothing
            ElementNode = Nothing
            TextNode = Nothing
        End Sub


        Public Sub addCDATAParameter(ByRef xmlQDBRequest As XmlDocument, ByVal Name As String, ByVal Value As String)
            Dim Root As XmlElement
            Dim ElementNode As XmlNode
            Dim CDATANode As XmlNode

            Root = xmlQDBRequest.DocumentElement
            ElementNode = xmlQDBRequest.CreateNode(XmlNodeType.Element, Name, "")
            CDATANode = xmlQDBRequest.CreateNode(XmlNodeType.CDATA, "", "")
            CDATANode.InnerText = Value
            ElementNode.AppendChild(CDATANode)
            Root.AppendChild(ElementNode)
            Root = Nothing
            ElementNode = Nothing
            CDATANode = Nothing
        End Sub

        Public Sub addFieldParameter(ByRef xmlQDBRequest As XmlDocument, ByVal attrName As String, ByVal Name As String, ByVal Value As Object)
            Dim Root As XmlElement
            Dim ElementNode As XmlNode
            Dim TextNode As XmlNode
            Dim attrField As XmlAttribute
            Dim attrFileName As XmlAttribute


            Root = xmlQDBRequest.DocumentElement
            ElementNode = xmlQDBRequest.CreateNode(XmlNodeType.Element, "field", "")
            attrField = xmlQDBRequest.CreateAttribute(attrName)
            attrField.Value = Name
            Call ElementNode.Attributes.SetNamedItem(attrField)


            If TypeName(Value) = "FileStream" Then
                attrFileName = xmlQDBRequest.CreateAttribute("filename")
                attrFileName.Value = DirectCast(Value, FileStream).Name
                Call ElementNode.Attributes.SetNamedItem(attrFileName)
            End If

            TextNode = xmlQDBRequest.CreateNode(XmlNodeType.Text, "", "")
            If TypeName(Value) = "FileStream" Then
                TextNode.InnerText = fileEncode64(DirectCast(Value, FileStream))
            Else
                TextNode.InnerText = CStr(Value)
            End If
            ElementNode.AppendChild(TextNode)

            Root.AppendChild(ElementNode)
            Root = Nothing
            ElementNode = Nothing
            attrField = Nothing
            TextNode = Nothing
        End Sub
        Function int64ToDate(ByVal int64 As String) As Date
            int64ToDate = Date.FromOADate(DAYS_BETWEEN_JAVASCRIPT_AND_MICROSOFT_DATE_REFERENCES + int64toDateCommon(int64))
        End Function
        Private Function int64toDateCommon(ByVal int64 As String) As Double
            If int64 = "" Then
                int64toDateCommon = vbNull
                Exit Function
            End If
            Dim dblTemp As Double
            dblTemp = makeDouble(int64)
            If dblTemp <= -59011459200001.0# Then
                int64toDateCommon = -59011459200000.0#
            ElseIf dblTemp > 255611376000000.0# Then
                int64toDateCommon = 255611376000000.0#
            Else
                int64toDateCommon = (dblTemp / MILLISECONDS_IN_A_DAY)
            End If
        End Function

        Function int64ToDuration(ByVal int64 As String) As Date
            int64ToDuration = Date.FromOADate(int64toDateCommon(int64))
        End Function

        Function makeAlphaNumLowerCase(ByVal strString As String) As String
            Dim i As Integer
            Dim chrString As String

            makeAlphaNumLowerCase = ""
            For i = 1 To Len(strString)
                chrString = Mid(strString, i, 1)
                If System.Char.IsLetterOrDigit(chrString, 0) Then
                    makeAlphaNumLowerCase = makeAlphaNumLowerCase & chrString
                Else
                    makeAlphaNumLowerCase = makeAlphaNumLowerCase & "_"
                End If
            Next i
            makeAlphaNumLowerCase = LCase(makeAlphaNumLowerCase)
        End Function
        Public Sub setGMTOffset(ByVal offsetHours As Single)
            GMTOffset = offsetHours
        End Sub
        Public Sub setServer(ByVal strHost As String, ByVal HTTPS As Boolean)
            If strHost <> "" Then
                QDBHost = strHost
                useHTTPS = HTTPS
            Else
                QDBHost = ""
                useHTTPS = True
            End If
        End Sub

        Public Function getDBLastModified(ByVal dbid As String) As Date
            Dim qdbResponse As New XmlDocument
            Dim strInt64Time As String

            qdbResponse = GetDBInfo(dbid)
            strInt64Time = qdbResponse.DocumentElement.SelectSingleNode("/*/lastRecModTime").InnerText
            If Left(strInt64Time, 1) = "-" Then
                getDBLastModified = #1/1/1970#
            Else
                getDBLastModified = int64ToDate(qdbResponse.DocumentElement.SelectSingleNode("/*/lastRecModTime").InnerText)
            End If
        End Function

        Function getCompleteCSVSnapshot(ByVal dbid As String) As String
            Dim FieldNodeList As XmlNodeList
            Dim xmlNode As XmlNode
            Dim qdbResponse As New XmlDocument
            Dim clist As String = ""

            qdbResponse = GetSchema(dbid)
            FieldNodeList = qdbResponse.DocumentElement.SelectNodes("/*/table/fields/field/@id")
            For Each xmlNode In FieldNodeList
                clist = clist + xmlNode.InnerText & "."
            Next xmlNode
            getCompleteCSVSnapshot = GenResultsTable(dbid, "{'0'.CT.''}", clist, "", "csv")
        End Function
        Function getRecordAsArray(ByVal dbid As String, ByVal clist As String, ByVal ridFID As String, ByVal rid As String, ByRef QDBRecord(,) As Object) As String
            Dim xmlQDBResponse As New XmlDocument
            xmlQDBResponse = DoQuery(dbid, "{'" & ridFID & "'.EX.'" & rid & "'", clist, "", "")
            Dim strFieldValue As String
            Dim i As Integer
            Dim FieldNodeList As XmlNodeList
            Dim FieldDefNodeList As XmlNodeList

            FieldDefNodeList = xmlQDBResponse.DocumentElement.SelectNodes("/*/table/fields/field")
            FieldNodeList = xmlQDBResponse.DocumentElement.SelectNodes("/*/table/records/record/f")

            ReDim QDBRecord(1, FieldNodeList.Count - 1)
            For i = 0 To FieldNodeList.Count - 1
                QDBRecord(0, i) = FieldNodeList(i).SelectSingleNode("@id").InnerText
                strFieldValue = FieldNodeList(i).SelectSingleNode(".").InnerText

                Select Case FieldDefNodeList(i).SelectSingleNode("@base_type").InnerText
                    Case "float"
                        If strFieldValue <> "" Then
                            QDBRecord(1, i) = makeDouble(strFieldValue)
                        End If
                    Case "text"
                        QDBRecord(1, i) = Replace(strFieldValue, Chr(10), vbCrLf)
                    Case "bool"
                        QDBRecord(1, i) = CBool(strFieldValue)
                    Case "int64"
                        If strFieldValue <> "" Then
                            If FieldDefNodeList(i).SelectSingleNode("@field_type").InnerText <> "date" Then
                                QDBRecord(1, i) = CDbl(strFieldValue) / MILLISECONDS_IN_A_DAY
                            Else
                                QDBRecord(1, i) = int64ToDate(strFieldValue)
                            End If
                        End If
                    Case "int32"
                        If FieldDefNodeList(i).SelectSingleNode("@field_type").InnerText = "userid" Then
                            On Error Resume Next
                            Dim tempLong As Integer
                            tempLong = CInt(strFieldValue)
                            If Err.Number = 0 Then
                                QDBRecord(1, i) = xmlQDBResponse.SelectSingleNode("/*/table/lusers/luser[@id='" & strFieldValue & "']").InnerText
                            Else
                                QDBRecord(1, i) = strFieldValue
                            End If
                            On Error GoTo 0
                        Else
                            QDBRecord(1, i) = CLng(strFieldValue)
                        End If
                End Select
            Next i
            getRecordAsArray = xmlQDBResponse.DocumentElement.SelectSingleNode("/*/table/records/record/f[@id=/*/table/fields/field[@role='modified']/@id]").InnerText
        End Function
        Public Function makeDouble(ByVal strString As String) As Double
            Dim i As Integer
            Dim chrString As String
            Dim strChar As String
            Dim resultString As String

            On Error Resume Next
            makeDouble = CDbl(strString)
            If Err.Number = 0 Then
                Exit Function
            End If
            On Error GoTo 0
            resultString = ""
            For i = 1 To Len(strString)
                strChar = Mid(strString, i, 1)
                If (((Not System.Char.IsLetter(strChar, 0)) And System.Char.IsLetterOrDigit(strChar, 0)) Or strChar = "." Or strChar = "-") Then
                    resultString = resultString & strChar
                End If
            Next i
            On Error Resume Next
            makeDouble = CDbl(resultString)
            Exit Function
        End Function


        Function fileEncode64(ByVal fileToUpload As FileStream) As String
            Dim triplicate As Integer
            Dim i As Integer
            Dim outputText As String
            Dim fileLength As Integer
            Dim fileTriads As Integer
            Dim firstByte(0) As Byte
            Dim secondByte(0) As Byte
            Dim thirdByte(0) As Byte
            Dim fileRemainder As Integer

            fileLength = CInt(fileToUpload.Length)
            fileRemainder = CInt(fileLength Mod 3)
            fileTriads = fileLength \ 3
            If fileRemainder > 0 Then
                outputText = Space((fileTriads + 1) * 4)
            Else
                outputText = Space(fileTriads * 4)
            End If


            For i = 0 To fileTriads - 1             ' loop through octets
                'build 24 bit triplicate
                fileToUpload.Read(firstByte, 0, 1)
                fileToUpload.Read(secondByte, 0, 1)
                fileToUpload.Read(thirdByte, 0, 1)

                triplicate = (CInt(firstByte(0)) * 65536) + (CInt(secondByte(0)) * CInt(256)) + CInt(thirdByte(0))
                'extract four 6 bit quartets from triplicate
                Mid(outputText, (i * 4) + 1) = Mid(Map64, (triplicate \ 262144) + 1, 1) & Mid(Map64, ((triplicate And 258048) \ 4096) + 1, 1) & Mid(Map64, ((triplicate And 4032) \ 64) + 1, 1) & Mid(Map64, (triplicate And 63) + 1, 1)
            Next                                                    ' next octet
            Select Case fileRemainder
                Case 1
                    fileToUpload.Read(firstByte, 0, 1)
                    triplicate = (firstByte(0) * 65536)
                    Mid(outputText, (i * 4) + 1) = Mid(Map64, (triplicate \ 262144) + 1, 1) & Mid(Map64, ((triplicate And 258048) \ 4096) + 1, 1) & "="
                Case 2
                    fileToUpload.Read(firstByte, 0, 1)
                    fileToUpload.Read(secondByte, 0, 1)
                    triplicate = (firstByte(0) * 65536) + (secondByte(0) * 256)
                    Mid(outputText, (i * 4) + 1) = Mid(Map64, (triplicate \ 262144) + 1, 1) & Mid(Map64, ((triplicate And 258048) \ 4096) + 1, 1) & Mid(Map64, ((triplicate And 4032) \ 64) + 1, 1) & "="
            End Select
            fileEncode64 = outputText
        End Function

        Function makeValidFilename(ByVal strString As String) As String
            Dim i As Integer
            Dim byteChar As String
            makeValidFilename = ""
            For i = 1 To Len(strString)
                byteChar = Mid(strString, i, 1)
                If byteChar = "\" Or byteChar = "/" Or
                   byteChar = ":" Or byteChar = "*" Or
                   Asc(byteChar) = 63 Or byteChar = """" Or
                   byteChar = "<" Or byteChar = ">" Or
                   byteChar = "|" Or byteChar = "'" _
                Then
                    makeValidFilename = makeValidFilename & "_"
                Else
                    makeValidFilename = makeValidFilename + byteChar
                End If
            Next i
        End Function
        Public Function getServerStatus() As XmlDocument
            Dim xmlQDBRequest As XmlDocument
            xmlQDBRequest = InitXMLRequest()
            getServerStatus = APIXMLPost("main", "API_OBStatus", xmlQDBRequest, useHTTPS)
        End Function
        Public Function GetNumRecords(ByVal dbid As String) As Integer
            Dim xmlQDBRequest As XmlDocument

            xmlQDBRequest = InitXMLRequest()
            GetNumRecords = CInt(APIXMLPost(dbid, "API_GetNumRecords", xmlQDBRequest, useHTTPS).SelectSingleNode("/*/num_records").InnerText)
        End Function

        Public Function GetDBPage(ByVal dbid As String, ByVal page As String) As String
            Dim xmlQDBRequest As XmlDocument

            xmlQDBRequest = InitXMLRequest()
            If CStr(Val(page)) = page Then
                addParameter(xmlQDBRequest, "pageid", page)
            Else
                addParameter(xmlQDBRequest, "pagename", page)
            End If
            GetDBPage = APIXMLPost(dbid, "API_GetDBPage", xmlQDBRequest, useHTTPS).SelectSingleNode("/*/pagebody").InnerText
        End Function

        Public Function AddReplaceDBPage(ByVal dbid As String, ByVal page As String, ByVal pagetype As String, ByVal pagebody As String) As XmlDocument
            Dim xmlQDBRequest As XmlDocument

            xmlQDBRequest = InitXMLRequest()
            If CStr(Val(page)) = page Then
                addParameter(xmlQDBRequest, "pageid", page)
            Else
                addParameter(xmlQDBRequest, "pagename", page)
            End If
            addParameter(xmlQDBRequest, "pagetype", pagetype)
            addParameter(xmlQDBRequest, "pagebody", pagebody)
            AddReplaceDBPage = APIXMLPost(dbid, "API_AddReplaceDBPage", xmlQDBRequest, useHTTPS)
        End Function

        Public Function ListDBPages(ByVal dbid As String) As XmlDocument
            Dim xmlQDBRequest As XmlDocument

            xmlQDBRequest = InitXMLRequest()
            ListDBPages = APIXMLPost(dbid, "API_ListDBPages", xmlQDBRequest, useHTTPS)
        End Function
        Public Function PurgeRecords(ByVal dbid As String, ByVal query As String) As XmlDocument
            Dim xmlQDBRequest As XmlDocument
            xmlQDBRequest = InitXMLRequest()
            If Left(query, 1) = "{" And Right(query, 1) = "}" Then
                addParameter(xmlQDBRequest, "query", query)
            ElseIf CStr(Val(query)) = query Then
                addParameter(xmlQDBRequest, "qid", query)
            Else
                addParameter(xmlQDBRequest, "qname", query)
            End If


            PurgeRecords = APIXMLPost(dbid, "API_PurgeRecords", xmlQDBRequest, useHTTPS)

        End Function
        Public Sub New(ByVal uT As String)
            userToken = uT
            GMTOffset = -7
        End Sub

        Public Function RenameApp(ByVal dbid As String, ByVal newappname As String) As Boolean
            Dim xmlQDBRequest As XmlDocument
            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "newappname", newappname)
            On Error GoTo exception
            Call APIXMLPost(dbid, "API_RenameApp", xmlQDBRequest, useHTTPS)
            RenameApp = True
            Exit Function
exception:
            RenameApp = False
        End Function

        Public Function GetDBvar(ByVal dbid As String, ByVal varname As String) As String
            Dim xmlQDBRequest As XmlDocument
            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "varname", varname)
            xmlQDBRequest = APIXMLPost(dbid, "API_GetDBvar", xmlQDBRequest, useHTTPS)
            GetDBvar = xmlQDBRequest.DocumentElement.SelectSingleNode("/*/value").InnerText
        End Function

        Public Function CreateTable(ByVal application_dbid As String, ByVal tname As String, ByVal pnoun As String) As String
            Dim xmlQDBRequest As XmlDocument
            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "pnoun", pnoun)
            addParameter(xmlQDBRequest, "tname", tname)
            xmlQDBRequest = APIXMLPost(application_dbid, "API_CreateTable", xmlQDBRequest, useHTTPS)
            CreateTable = xmlQDBRequest.DocumentElement.SelectSingleNode("/*/newdbid").InnerText
        End Function

        Public Function AddUserToRole(ByVal dbid As String, ByVal userid As String, ByVal roleid As String) As Boolean
            Dim xmlQDBRequest As XmlDocument
            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "userid", userid)
            addParameter(xmlQDBRequest, "roleid", roleid)
            On Error GoTo exception
            Call APIXMLPost(dbid, "API_AddUserToRole", xmlQDBRequest, useHTTPS)
            AddUserToRole = True
            Exit Function
exception:
            AddUserToRole = False
        End Function

        Public Function GetOneTimeTicket() As String
            Dim xmlQDBRequest As XmlDocument
            xmlQDBRequest = InitXMLRequest()
            xmlQDBRequest = APIXMLPost("main", "API_GetOneTimeTicket", xmlQDBRequest, useHTTPS)
            GetOneTimeTicket = xmlQDBRequest.DocumentElement.SelectSingleNode("/*/ticket").InnerText
        End Function

        Public Function FieldRemoveChoices(ByVal dbid As String, ByVal fid As String, ByVal ParamArray NameValues() As Object) As Integer
            Dim xmlQDBRequest As XmlDocument
            Dim firstfield As Integer
            Dim lastfield As Integer
            Dim i As Integer

            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "fid", fid)

            lastfield = UBound(NameValues)
            firstfield = LBound(NameValues)
            For i = firstfield To lastfield
                addParameter(xmlQDBRequest, "choice", CStr(NameValues(i)))
            Next i

            xmlQDBRequest = APIXMLPost(dbid, "API_FieldRemoveChoices", xmlQDBRequest, useHTTPS)
            FieldRemoveChoices = CInt(xmlQDBRequest.DocumentElement.SelectSingleNode("/*/numremoved").InnerText)
        End Function

        Public Function DeleteField(ByVal dbid As String, ByVal fid As String) As Boolean
            Dim xmlQDBRequest As XmlDocument
            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "fid", fid)
            On Error GoTo exception
            Call APIXMLPost(dbid, "API_DeleteField", xmlQDBRequest, useHTTPS)
            DeleteField = True
            Exit Function
exception:
            DeleteField = False
        End Function

        Public Function GenAddRecordForm(ByVal dbid As String, ByRef fieldValues As Hashtable) As String
            Dim xmlQDBRequest As XmlDocument
            xmlQDBRequest = InitXMLRequest()
            Dim fieldValueEnumerator As IDictionaryEnumerator = fieldValues.GetEnumerator()
            While fieldValueEnumerator.MoveNext()
                addParameterWithAttribute(xmlQDBRequest, "field", "name", fieldValueEnumerator.Key.ToString, fieldValueEnumerator.Value.ToString())
            End While
            GenAddRecordForm = APIHTMLPost(dbid, "API_GenAddRecordForm", xmlQDBRequest, useHTTPS)
        End Function

        Public Function ChangeUserRole(ByVal dbid As String, ByVal userid As String, ByVal roleid As String, ByVal newroleid As String) As Boolean
            Dim xmlQDBRequest As XmlDocument
            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "userid", userid)
            addParameter(xmlQDBRequest, "roleid", roleid)
            addParameter(xmlQDBRequest, "newroleid", newroleid)
            On Error GoTo exception
            Call APIXMLPost(dbid, "API_ChangeUserRole", xmlQDBRequest, useHTTPS)
            ChangeUserRole = True
            Exit Function
exception:
            ChangeUserRole = False
        End Function

        Public Function SetDBvar(ByVal dbid As String, ByVal varname As String, ByVal value As String) As Boolean
            Dim xmlQDBRequest As XmlDocument
            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "varname", varname)
            addParameter(xmlQDBRequest, "value", value)
            On Error GoTo exception
            Call APIXMLPost(dbid, "API_SetDBvar", xmlQDBRequest, useHTTPS)
            SetDBvar = True
            Exit Function
exception:
            SetDBvar = False
        End Function

        Public Function ProvisionUser(ByVal dbid As String, ByVal roleid As String, ByVal email As String, ByVal fname As String, ByVal lname As String) As String
            Dim xmlQDBRequest As XmlDocument
            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "roleid", roleid)
            addParameter(xmlQDBRequest, "email", email)
            addParameter(xmlQDBRequest, "fname", fname)
            addParameter(xmlQDBRequest, "lname", lname)
            xmlQDBRequest = APIXMLPost(dbid, "API_ProvisionUser", xmlQDBRequest, useHTTPS)
            ProvisionUser = xmlQDBRequest.DocumentElement.SelectSingleNode("/*/userid").InnerText
        End Function

        Public Function GetRecordInfo(ByVal dbid As String, ByVal rid As String) As XmlDocument
            Dim xmlQDBRequest As XmlDocument
            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "rid", rid)
            GetRecordInfo = APIXMLPost(dbid, "API_GetRecordInfo", xmlQDBRequest, useHTTPS)
        End Function

        Public Function UserRoles(ByVal dbid As String) As XmlDocument
            Dim xmlQDBRequest As XmlDocument
            xmlQDBRequest = InitXMLRequest()
            UserRoles = APIXMLPost(dbid, "API_UserRoles", xmlQDBRequest, useHTTPS)
        End Function

        Public Function GetUserRole(ByVal dbid As String, ByVal userid As String) As XmlDocument
            Dim xmlQDBRequest As XmlDocument
            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "userid", userid)
            GetUserRole = APIXMLPost(dbid, "API_GetUserRole", xmlQDBRequest, useHTTPS)
        End Function

        Public Function RemoveUserFromRole(ByVal dbid As String, ByVal userid As String, ByVal roleid As String) As Boolean
            Dim xmlQDBRequest As XmlDocument
            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "userid", userid)
            addParameter(xmlQDBRequest, "roleid", roleid)
            On Error GoTo exception
            Call APIXMLPost(dbid, "API_RemoveUserFromRole", xmlQDBRequest, useHTTPS)
            RemoveUserFromRole = True
            Exit Function
exception:
            RemoveUserFromRole = False
        End Function

        Public Function GetRecordAsHTML(ByVal dbid As String, ByVal rid As String) As String
            Dim xmlQDBRequest As XmlDocument
            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "rid", rid)
            GetRecordAsHTML = APIHTMLPost(dbid, "API_GetRecordAsHTML", xmlQDBRequest, useHTTPS)
        End Function

        Public Function SendInvitation(ByVal dbid As String, ByVal userid As String) As Boolean
            Dim xmlQDBRequest As XmlDocument
            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "userid", userid)
            On Error GoTo exception
            Call APIXMLPost(dbid, "API_SendInvitation", xmlQDBRequest, useHTTPS)
            SendInvitation = True
            Exit Function
exception:
            SendInvitation = True
        End Function

        Public Function GetUserInfo(ByVal email As String) As XmlDocument
            Dim xmlQDBRequest As XmlDocument
            xmlQDBRequest = InitXMLRequest()
            addParameter(xmlQDBRequest, "email", email)
            GetUserInfo = APIXMLPost("main", "API_GetUserInfo", xmlQDBRequest, useHTTPS)
        End Function

        Public Function GetRoleInfo(ByVal dbid As String) As XmlDocument
            Dim xmlQDBRequest As XmlDocument
            xmlQDBRequest = InitXMLRequest()
            GetRoleInfo = APIXMLPost(dbid, "API_GetRoleInfo", xmlQDBRequest, useHTTPS)
        End Function

    End Class
End Module
