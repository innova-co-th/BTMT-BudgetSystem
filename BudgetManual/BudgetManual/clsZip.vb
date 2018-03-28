'   Database backup utility:
'   ========================
'   Copyright (C) 2007  Shabdar Ghata 
'   Email : ghata2002@gmail.com
'   URL : http://www.shabdar.org

'   This program is free software: you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or
'   (at your option) any later version.

'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.

'   You should have received a copy of the GNU General Public License
'   along with this program.  If not, see <http://www.gnu.org/licenses/>.

'   This program comes with ABSOLUTELY NO WARRANTY.

Imports System.IO
Imports ICSharpCode.SharpZipLib.Checksums
Imports ICSharpCode.SharpZipLib.Zip
Imports ICSharpCode.SharpZipLib.GZip
Public Class clsZip
    Public Shared Sub ZipIT(ByVal sSourceDir As String, ByVal sFileName As String)

        Dim astrFileNames() As String = Directory.GetFiles(sSourceDir)
        Dim objCrc32 As New Crc32()
        Dim strmZipOutputStream As ZipOutputStream

        strmZipOutputStream = New ZipOutputStream(File.Create((sFileName)))
        strmZipOutputStream.SetLevel(6)

        REM Compression Level: 0-9
        REM 0: no(Compression)
        REM 9: maximum compression

        Dim strFile As String

        For Each strFile In astrFileNames

            Dim strmFile As FileStream = File.OpenRead(strFile)
            Dim abyBuffer(CInt(strmFile.Length - 1)) As Byte

            strmFile.Read(abyBuffer, 0, abyBuffer.Length)

            Dim objZipEntry As ZipEntry = New ZipEntry(GetFileNameFromFullPath(strFile))

            objZipEntry.DateTime = DateTime.Now
            objZipEntry.Size = strmFile.Length

            strmFile.Close()
            objCrc32.Reset()
            objCrc32.Update(abyBuffer)
            objZipEntry.Crc = objCrc32.Value
            strmZipOutputStream.PutNextEntry(objZipEntry)
            strmZipOutputStream.Write(abyBuffer, 0, abyBuffer.Length)

        Next
        strmZipOutputStream.Finish()
        strmZipOutputStream.Close()
    End Sub
    Public Shared Function GetFileNameFromFullPath(ByVal sPath As String) As String
        Dim n As Integer = InStrRev(sPath, "\")
        If n > 0 Then
            Return Trim(Mid(sPath, n + 1, Len(sPath)))
        End If
        Return sPath
    End Function

    Public Shared Sub UnzipIT(ByVal sFileName As String, ByVal sDirName As String)
        Dim s As New ZipInputStream(File.OpenRead(sFileName))
        Dim theEntry As ZipEntry
        theEntry = s.GetNextEntry()
        Do While Not (IsNothing(theEntry))
            Dim directoryName As String = Path.GetDirectoryName(theEntry.Name)
            Dim fileName As String = Path.GetFileName(theEntry.Name)
            Directory.CreateDirectory(sDirName)
            Dim sWriter As FileStream
            sWriter = File.Create(sDirName + fileName)
            Dim size As Integer = 2048
            Dim data(2048) As Byte
            Do While (True)
                size = s.Read(data, 0, 2049)
                If (size > 0) Then
                    sWriter.Write(data, 0, size)
                Else
                    GoTo ex1
                End If
            Loop
EX1:
            sWriter.Close()
            theEntry = s.GetNextEntry
        Loop
        s.Close()
    End Sub
End Class
