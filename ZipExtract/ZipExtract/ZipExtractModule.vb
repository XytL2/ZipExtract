Imports System.IO
Imports System.IO.Compression
Imports System.Xml

Module ZipExtractModule


  Private Sub ExtractAndLoadXml(ByVal zipFilePath As String, ByVal doc As XmlDocument)
    Using fs As FileStream = File.OpenRead(zipFilePath)
      ExtractAndLoadXml(fs, doc)
    End Using
  End Sub

  Private Sub ExtractAndLoadXml(ByVal fs As FileStream, ByVal doc As XmlDocument)
    Dim extract = "C:\Extracted"
    Dim Archive As New ZipArchive(fs)

    For Each elementInsideZip As ZipArchiveEntry In Archive.Entries
      Dim ZipArchiveName As String = elementInsideZip.Name
      If Not ZipArchiveName.Contains(".zip") AndAlso ZipArchiveName <> "" Then
        Dim sdate = elementInsideZip.LastWriteTime.ToString
        Dim zipStream As Stream = elementInsideZip.Open()
        Dim extractedfile As FileStream = File.OpenWrite(Path.Combine(extract, ZipArchiveName))
        zipStream.CopyTo(extractedfile)
        Console.WriteLine("Moved " & ZipArchiveName & " to " & Path.Combine(extract, ZipArchiveName))
        extractedfile.Flush()
        extractedfile.Close()
        'Exit For
      ElseIf ZipArchiveName.Contains(".zip") Then
        Dim zipStream As Stream = elementInsideZip.Open()

        Dim zipFileExtractPath As String = Path.GetTempFileName()
        Dim extractedZipFile As FileStream = File.OpenWrite(zipFileExtractPath)
        zipStream.CopyTo(extractedZipFile)
        extractedZipFile.Flush()
        extractedZipFile.Close()
        Try
          ExtractAndLoadXml(zipFileExtractPath, doc)
        Finally
          File.Delete(zipFileExtractPath)
        End Try
      End If
    Next
  End Sub

  Public Sub Main()
    Dim location As String = "C:\Users\johnr\Desktop\FolderstructTop.zip"
    Dim xmlDocument As XmlDocument = New XmlDocument()
    ExtractAndLoadXml(location, xmlDocument)
    Console.ReadLine()
  End Sub
  'Public Sub main()

  '  Dim start = "C:\Users\johnr\Desktop\FolderstructTop.zip"
  '  'Dim test = ZipFile.Open(start, ZipArchiveMode.Read)
  '  Dim extract = "C:\Users\johnr\Desktop"
  '  Dim top As New MemoryStream()
  '  Dim stream As New MemoryStream

  '  Dim Archive As New ZipArchive(fs)
  '  m Archive = ZipFile.OpenRead(start)
  '  'Using filestream As FileStream = System.IO.File.Open(start, System.IO.FileMode.Open)

  '  Dim parentArchive = New ZipArchive(FileStream)
  '  For Each a In parentArchive.Entries
  '    If a.Length > 0 Then
  '      Dim zipStream As Stream = fs.GetInputStream(elementInsideZip)
  '      Dim filename = Path.GetFileName(a.FullName)
  '      Dim entry = parentArchive.GetEntry(filename)
  '      Dim childArchive = New ZipArchive(entry.Open())
  '      For Each e In childArchive.Entries

  '        Console.WriteLine(e.Name)
  '      Next
  '    End If



  '  Next
  'End Using
  'End Sub


  'Using archive As ZipArchive = ZipFile.OpenRead(start)
  '  For Each entry As ZipArchiveEntry In archive.Entries
  '    If entry.FullName.EndsWith(".png", StringComparison.InvariantCultureIgnoreCase) Then
  '      entry.ExtractToFile(Path.Combine(extract, entry.FullName))
  '    End If
  '  Next



  'End Using




  'For Each zip In test.Entries
  '  If zip.Length <> 0 Then
  '    'For Each zip2 In zip.entries
  '    Dim filename = Path.GetFileName(zip.FullName)
  '    'Next
  '  Else

  '    '  Dim filename = Path.GetFileName(zip.FullName)

  '    'Dim filename = zip.FullName.ToString.Substring(zip.FullName.ToString.LastIndexOf("/"), zip.FullName.Length - 1)
  '    ' zip.ExtractToFile(extract + filename)
  '    'ZipFile.ExtractToDirectory(start & zip.FullName.ToString, extract)
  '    'Console.WriteLine("Unzipping " & zip.Name & " to " & extract)
  '    'zip.ExtractToFile(extract)
  '  End If
  'Next


  'Sub unpackAll(pat As String)
  '  If Not Directory.Exists(pat) Then Exit Sub
  '  For Each zfn In Directory.GetFiles(pat, "*.zip")
  '    Try
  '      ZipFile.ExtractToDirectory(zfn, pat)
  '    Catch ex As Exception
  '    End Try
  '  Next
  'End Sub
End Module
