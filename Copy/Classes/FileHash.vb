Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Public Class FileHash


	Private m_file As IO.FileInfo
	Private m_fileHash As Byte()

	Public Property File As IO.FileInfo
		Get
			Return m_file
		End Get
		Set(value As IO.FileInfo)
			m_file = value
		End Set
	End Property

	Public Property FileHash As Byte()
		Get
			Return m_fileHash
		End Get
		Set(value As Byte())
			m_fileHash = value
		End Set
	End Property

	Public Sub New()
		m_file = Nothing
		m_fileHash = Nothing
	End Sub


    Public Sub New(ByRef fFile As IO.FileInfo)

        Try
            Init(fFile)
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Sub New(ByRef strPath As String)
        Try
            Init(New FileInfo(strPath))
        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    Private Sub Init(ByRef fFIle As IO.FileInfo)
        Try
            If fFIle IsNot Nothing Then
                m_file = fFIle
            Else
                Throw New NullReferenceException("File cannot be null: FileHash.New(fFile, abytFileHash)")
            End If

            m_fileHash = MD5CryptoServiceProvider.Create.ComputeHash(IO.File.ReadAllBytes(m_file.FullName))

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Overrides Function ToString() As String
		Dim strReturn As String = String.Empty
		Try

			strReturn = ByteArrayToString(m_fileHash)

		Catch ex As Exception
			Throw ex
		End Try

		Return strReturn

	End Function

	Public Shared Function ByteArrayToString(ByVal arrInput() As Byte) As String
		Dim i As Integer
		Dim sOutput As New StringBuilder(arrInput.Length)
		For i = 0 To arrInput.Length - 1
			sOutput.Append(arrInput(i).ToString("X2"))
		Next
		Return sOutput.ToString
	End Function

	Public Shared Function GetFileHash(ByVal fFile As FileInfo) As Byte()
		Dim byt As Byte() = Nothing

		byt = MD5CryptoServiceProvider.Create.ComputeHash(IO.File.ReadAllBytes(fFile.FullName))

		Return byt
	End Function

	Public Shared Function GetFileHashString(ByVal fFile As FileInfo) As String
		Dim strReturn As String = String.Empty

		strReturn = ByteArrayToString(GetFileHash(fFile))

		Return strReturn
	End Function

End Class
