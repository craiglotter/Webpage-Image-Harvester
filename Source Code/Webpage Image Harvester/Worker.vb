Imports System.Net
Imports System.IO
Imports System.Text


Public Class Worker

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.

    Public WorkerThread As System.Threading.Thread

    Public download As Download_Item

    Public url As String
    Public download_item_queue As ArrayList
    Public acceptable_items As ArrayList

    Public proxy As String
    Public savepath As String
    Public filehandle As Integer

    Public result As String = ""

    Public Event WorkerSnapShotTaken(ByVal result As String)
    Public Event WorkerStringMessage(ByVal message As String, ByVal labelname As String)

    Public Event WorkerError(ByVal Message As Exception)
    Public Event WorkerComplete(ByVal Result As String)
    Public Event WorkerProgress(ByVal value As Integer, ByVal read As Long, ByVal toread As Long, ByVal unit As String, ByVal bytesread As Long, ByVal bytestoread As Long)



#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)

    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        download_item_queue = New ArrayList
        acceptable_items = New ArrayList
    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message("The Application encountered the following problem: " & vbCrLf & identifier_msg & ":" & ex.ToString)
                Display_Message1.ShowDialog()
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ":" & ex.ToString)
                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception
            MsgBox("An error occurred in Webpage Image Harvester's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub



    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.


            Select Case threadNumber
                Case 1
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerExecute)
                    WorkerThread.Start()
                Case 2
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerTakeSnapShot)
                    WorkerThread.Start()
            End Select
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerExecute()
        Try

            Dim filenametodownload As String
            If download.filesavelevel = 0 Then
                Dim dinfo As DirectoryInfo = New DirectoryInfo(savepath)
                If dinfo.Exists = False Then
                    dinfo.Create()
                End If
                filenametodownload = savepath & download.filename
            Else
                Dim filler As String
                filler = download.filesavelevel.ToString
                While filler.Length < 3
                    filler = "0" & filler
                End While
                Dim dinfo As DirectoryInfo = New DirectoryInfo(savepath & filler & "\")
                If dinfo.Exists = False Then
                    dinfo.Create()
                End If

                filenametodownload = savepath & filler & "\" & download.filename
            End If
            Dim finfo As FileInfo = New FileInfo(filenametodownload)
            Dim dodownload As Boolean
            dodownload = True
            result = "Downloaded"
            If finfo.Exists = True Then
                If filehandle = 1 Then
                    finfo.Delete()
                    dodownload = True
                    result = "Downloaded"
                End If
                If filehandle = 2 Then
                    filenametodownload = filenametodownload.Insert(filenametodownload.Length - 4, "R1_" & Now.Ticks)
                    dodownload = True
                    result = "Downloaded (Renamed)"
                End If
                If filehandle = 3 Then
                    dodownload = False
                    result = "Ignored"
                End If
            End If
            If dodownload = True Then

                Dim wr As HttpWebRequest = CType(HttpWebRequest.Create(download.fileurl), HttpWebRequest)
                If proxy.Length > 0 Then
                    ' txtStatus_Message("Importing Proxy Information")
                    Dim proxyObject As New System.Net.WebProxy(proxy, True)
                    wr.Proxy = proxyObject
                End If

                Dim ws As HttpWebResponse = CType(wr.GetResponse(), HttpWebResponse)
                Dim str As Stream = ws.GetResponseStream()
                'MsgBox(ws.ContentLength)
                Dim inBuf(3000) As Byte


                Dim bytesToRead As Long
                Dim bytesRead As Long
                Dim totalbytesread As Long = 0
                Dim fstr As New FileStream(filenametodownload, FileMode.Create, FileAccess.Write)
                Dim displayunit As Integer
                Dim displayunittext As String
                If ws.ContentLength <= 1024 Then
                    displayunit = 1
                    displayunittext = "bytes"
                End If
                If ws.ContentLength > 1024 And ws.ContentLength <= (1024 * 1024) Then
                    displayunit = 1024
                    displayunittext = "KB"
                End If
                If ws.ContentLength > (1024 * 1024) And ws.ContentLength <= (1024 * 1024 * 1024) Then
                    displayunit = 1024 * 1024
                    displayunittext = "MB"
                End If
                If ws.ContentLength > (1024 * 1024 * 1024) Then
                    displayunit = 1024 * 1024 * 1024
                    displayunittext = "GB"
                End If
                Try
                    While fstr.Length < ws.ContentLength

                        If ws.ContentLength - fstr.Length >= 3000 Then
                            bytesToRead = 3000
                            ReDim inBuf(bytesToRead)
                        Else
                            bytesToRead = ws.ContentLength - fstr.Length
                            ReDim inBuf(bytesToRead)
                        End If

                        bytesRead = str.Read(inBuf, 0, bytesToRead)



                        'bytesToRead -= n
                        totalbytesread = totalbytesread + bytesRead
                        RaiseEvent WorkerProgress(Math.Round(((totalbytesread) / (ws.ContentLength) * 100), 0) * 2, Math.Round(totalbytesread / displayunit, 0), Math.Round(ws.ContentLength / displayunit, 0), displayunittext, totalbytesread, ws.ContentLength)
                        fstr.Write(inBuf, 0, bytesRead)
                        'End While

                    End While

                    If fstr.Length <> ws.ContentLength Then
                        result = "Downloaded (Corrupt Download)"
                    End If
                Catch ex As Exception
                    result = "Download Failed"
                    Error_Handler(ex)
                End Try
                str.Close()
                fstr.Close()


            End If




        Catch ex As Exception
            result = "Download Failed"
            Error_Handler(ex)
        End Try

        Activity_Logger(download.fileurl & ": " & result)
        RaiseEvent WorkerComplete(result)
    End Sub

    Public Sub WorkerTakeSnapShot()

        Try
            Dim urlstring As String = ""
            url = url.Trim
            If url.EndsWith(";") Then
                url = url.Remove(url.Length - 1, 1)
            End If
            Dim stringarray As String() = url.Split(";")
            Dim savelevel As Integer = 0
            If stringarray.Length > 1 Then
                savelevel = 1
            End If
            Dim counturlstring As Integer
            counturlstring = 0
            RaiseEvent WorkerProgress(Math.Round((counturlstring / stringarray.Length * 100), 0) * 2, Math.Round(counturlstring, 0), Math.Round(stringarray.Length, 0), "URLs", 0, 0)
            For Each urlstring In stringarray
                counturlstring = counturlstring + 1
                Try

                    urlstring = urlstring.Trim
                    If urlstring.EndsWith("/") = False Then

                        If urlstring.LastIndexOf(".") = -1 Then
                            urlstring = urlstring & "/"
                        Else
                            If urlstring.LastIndexOf(".") < urlstring.LastIndexOf("/") Then
                                urlstring = urlstring & "/"
                            End If
                        End If
                    End If

                    Dim sServer As String
                    sServer = Trim(urlstring)

                    Dim errorhandled As Boolean = False

                    RaiseEvent WorkerStringMessage("Setting up HttpWebRequest Object for " & sServer, "txtStatus")


                    Dim HttpWReq As System.Net.HttpWebRequest = _
                              CType(System.Net.WebRequest.Create(sServer), System.Net.HttpWebRequest)


                    If proxy.Length > 0 Then
                        RaiseEvent WorkerStringMessage("Importing Proxy Information (" & proxy & ")", "txtStatus")

                        Dim proxyObject As New System.Net.WebProxy(proxy, True)
                        HttpWReq.Proxy = proxyObject
                    End If



                    RaiseEvent WorkerStringMessage("Getting HttpWebResponse from " & sServer, "txtStatus")
                    Dim HttpWResp As System.Net.HttpWebResponse = _
                       CType(HttpWReq.GetResponse(), System.Net.HttpWebResponse)

                    RaiseEvent WorkerStringMessage("Parsing HttpWebResponse from " & sServer & " for Valid HREF Tags", "txtStatus")

                    Dim streamer As System.IO.StreamReader = New System.IO.StreamReader(HttpWResp.GetResponseStream, System.Text.Encoding.ASCII, False, 512)

                    Dim stringtoanalyze As String
                    Dim substring As String
                    Dim acceptable As String
                    Dim addeditem As Boolean
                    While streamer.Peek() <> -1
                        If streamer.Peek <> -1 Then
                            stringtoanalyze = streamer.ReadLine.ToLower()
                        End If
                        If streamer.Peek <> -1 Then
                            stringtoanalyze = stringtoanalyze & " " & streamer.ReadLine.ToLower()
                        End If
                        If streamer.Peek <> -1 Then
                            stringtoanalyze = stringtoanalyze & " " & streamer.ReadLine.ToLower()
                        End If
                        If streamer.Peek <> -1 Then
                            stringtoanalyze = stringtoanalyze & " " & streamer.ReadLine.ToLower()
                        End If
                        'stringtoanalyze = streamer.ReadToEnd.ToLower()


                        RaiseEvent WorkerStringMessage(stringtoanalyze, "txtParse")

                        Try
                            While stringtoanalyze.IndexOf("src=") > 0
                                If stringtoanalyze.IndexOf("src=""javascript:") < 0 And stringtoanalyze.IndexOf("src=javascript:") < 0 And stringtoanalyze.IndexOf("src='javascript:") < 0 Then
                                    'MsgBox(stringtoanalyze)
                                    substring = stringtoanalyze.Substring(stringtoanalyze.IndexOf("src=") + 5, (stringtoanalyze.IndexOf(">", stringtoanalyze.IndexOf("src=") + 5)) - (stringtoanalyze.IndexOf("src=") + 5))
                                    'MsgBox(substring)
                                    stringtoanalyze = stringtoanalyze.Remove(0, (stringtoanalyze.IndexOf(">", stringtoanalyze.IndexOf("src=") + 5)) - 1)


                                    Dim hrefwithquotations As Boolean
                                    hrefwithquotations = False

                                    If substring.StartsWith("""") = True Or substring.StartsWith("'") = True Then
                                        substring = substring.Remove(0, 1)
                                        hrefwithquotations = True
                                    End If
                                    'MsgBox(substring)
                                    If hrefwithquotations = False Then
                                        If substring.IndexOf(" ") > -1 Then
                                            substring = substring.Remove(substring.IndexOf(" "), substring.Length - substring.IndexOf(" "))
                                        End If
                                    End If
                                    'MsgBox(substring)
                                    If substring.IndexOf("""") > -1 Then
                                        substring = substring.Remove(substring.IndexOf(""""), substring.Length - substring.IndexOf(""""))
                                    End If
                                    If substring.EndsWith("'") = True And hrefwithquotations = True Then
                                        substring = substring.Remove(substring.Length - 1, 1)
                                    End If
                                    If substring.EndsWith("/") Then
                                        substring = substring.Remove(substring.Length - 1, 1)
                                    End If
                                    'MsgBox(substring)
                                    Dim tempurlstring As String
                                    tempurlstring = urlstring

                                    If tempurlstring.LastIndexOf("?") > -1 Then
                                        'MsgBox(tempurlstring & "   " & tempurlstring.LastIndexOf("?") & "   " & tempurlstring.Length)
                                        tempurlstring = tempurlstring.Remove(tempurlstring.LastIndexOf("?"), tempurlstring.Length - tempurlstring.LastIndexOf("?"))
                                        'MsgBox(tempurlstring & "   " & tempurlstring.LastIndexOf("?") & "   " & tempurlstring.Length)
                                    End If

                                    If Not substring.StartsWith("http:") Then
                                        If substring.StartsWith("/") Then

                                            substring = tempurlstring.Substring(0, (tempurlstring.IndexOf("/", 7))) & substring
                                        Else

                                            substring = tempurlstring.Substring(0, (tempurlstring.LastIndexOf("/") + 1)) & substring
                                        End If
                                    End If
                                    addeditem = False
                                    For Each acceptable In acceptable_items

                                        If addeditem = False Then
                                            If substring.EndsWith(acceptable) Then
                                                Dim downitem As Download_Item = New Download_Item(substring.Replace("%5F", "_").Replace("%20", " ").Trim(), substring.Remove(0, substring.LastIndexOf("/") + 1).Replace("%5F", "_").Replace("%20", " "), download_item_queue.Count, True, savelevel)
                                                download_item_queue.Add(downitem)
                                                addeditem = True
                                            End If
                                        End If
                                    Next
                                Else
                                    stringtoanalyze = stringtoanalyze.Remove(stringtoanalyze.IndexOf("src=""javascript:"), 6)
                                End If
                            End While
                        Catch ex As Exception
                            Error_Handler(ex, stringtoanalyze)
                        End Try
                    End While


                    streamer.Close()

                    HttpWResp.Close()

                    RaiseEvent WorkerStringMessage("IMG Tag Parsing Complete", "txtStatus")
                Catch ex As Exception

                    Error_Handler(ex)
                End Try
                savelevel = savelevel + 1
                RaiseEvent WorkerProgress(Math.Round((counturlstring / stringarray.Length * 100), 0) * 2, Math.Round(counturlstring, 0), Math.Round(stringarray.Length, 0), "URLs", 0, 0)
            Next
            RaiseEvent WorkerStringMessage("IMG Tag Parsing Complete", "txtStatus")
        Catch ex As Exception

            Error_Handler(ex)
        End Try
        RaiseEvent WorkerSnapShotTaken("Complete")
    End Sub

    Private Function DosShellCommand(ByVal AppToRun As String) As String
        Dim s As String = ""
        Try
            Dim myProcess As Process = New Process

            myProcess.StartInfo.FileName = "cmd.exe"
            myProcess.StartInfo.UseShellExecute = False
            myProcess.StartInfo.CreateNoWindow = True
            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True
            myProcess.Start()
            Dim sIn As StreamWriter = myProcess.StandardInput
            sIn.AutoFlush = True

            Dim sOut As StreamReader = myProcess.StandardOutput
            Dim sErr As StreamReader = myProcess.StandardError
            sIn.Write(AppToRun & _
               System.Environment.NewLine)
            sIn.Write("exit" & System.Environment.NewLine)
            s = sOut.ReadToEnd()
            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If

            'MessageBox.Show("The 'dir' command window was closed at: " & myProcess.ExitTime & "." & System.Environment.NewLine & "Exit Code: " & myProcess.ExitCode)

            sIn.Close()
            sOut.Close()
            sErr.Close()
            myProcess.Close()
            'MessageBox.Show(s)
        Catch ex As Exception
            Error_Handler(ex)
        End Try
        Return s
    End Function



    Private Sub Activity_Logger(ByVal log As String)
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt", True)
            Try
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & log)
            Catch ex As Exception
                Error_Handler(ex, "Activity Logger")
            Finally
                filewriter.Flush()
                filewriter.Close()

            End Try
        Catch ex As Exception
            Error_Handler(ex, "Activity Logger")
        End Try
    End Sub





End Class
