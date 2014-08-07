Public Class Download_Item
    Public fileurl As String
    Public filename As String
    Public fileindex As Integer
    Public filedownload As Boolean
    Public filesavelevel As Integer

    Public Sub New()
        fileurl = ""
        filename = ""
        fileindex = ""
        filedownload = False
        filesavelevel = 0
    End Sub

    Public Sub New(ByVal url As String, ByVal name As String, ByVal index As String, ByVal download As Boolean, ByVal savelevel As Integer)
        fileurl = url
        filename = name
        fileindex = index
        filedownload = download
        filesavelevel = savelevel
    End Sub

End Class
