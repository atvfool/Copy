Imports System.IO
Imports System.Threading

Module Copy

	Private m_sourceHash As FileHash
	Private m_destHash As FileHash
	'Private m_lstSource As List(Of FileInfo)
	Private m_afi() As FileInfo
	Private m_lstCopied As List(Of String)
	Private m_strSourcePath As String
	Private m_strDestPath As String
	Private m_ccDefault As ConsoleColor
	Private m_CompareType As New CompareType
	Private Enum CompareType
		CopyNewer = 0
		CopyOlder
		CopyDiffHash
		CopyAll
	End Enum

	Sub Main()
		m_ccDefault = Console.ForegroundColor
		Dim args As String() = Environment.GetCommandLineArgs
		' Display help Menu and Exit?
		If args.Length = 1 OrElse args.Contains("-h") Then
			DisplayHelp()
#If DEBUG Then
			Console.Write("Press ENTER to exit...")
			Console.ReadKey()
#End If
			End
		End If

		' Get the Compare type
		If args.Contains("-C") Then
			For i As Integer = 0 To args.Length - 1
				If args(i) = "-C" Then
					m_CompareType = DirectCast([Enum].Parse(GetType(CompareType), args(i + 1)), CompareType)
				End If
			Next
		Else
			m_CompareType = CompareType.CopyAll
		End If
		m_strSourcePath = args(1)
		m_strDestPath = args(2)

		Dim tmp As String = String.Empty
		Dim intCount As Integer = 0
		Dim dtStart As DateTime = DateTime.Now
		Dim dtFinish As DateTime
		If IO.Directory.Exists(m_strSourcePath) Then
			LoadFilesArray(m_afi, New DirectoryInfo(m_strSourcePath))
			m_lstCopied = New List(Of String)

			Select Case m_CompareType
				Case CompareType.CopyNewer
					intCount = CopyNewer()
				Case CompareType.CopyAll
					intCount = CopyAll()
				Case CompareType.CopyOlder
					intCount = CopyOlder()
				Case CompareType.CopyDiffHash
					intCount = CopyHash()
			End Select

		End If
		dtFinish = DateTime.Now
		Dim ts As TimeSpan = dtFinish.Subtract(dtStart)
		WriteLogFile(intCount, ts)

		Console.WriteLine("--------------------------------------------------------------------------------")

		Console.WriteLine("Finished Copying " & intCount & " files")

		Console.WriteLine("It took " & Math.Floor(dtFinish.Subtract(dtStart).TotalMinutes) & " min and " & dtFinish.Subtract(dtStart).Seconds & " seconds")

		Console.WriteLine("--------------------------------------------------------------------------------")
#If DEBUG Then
		Console.Write("Press ENTER to exit...")
		Console.ReadKey()
#End If

	End Sub

	''' <summary>
	''' Checks if a file doesn't exists in the destination and copies or if it does exist, is the source newer, then copy.
	''' </summary>
	''' <returns></returns>
	Public Function CopyNewer() As Integer
		Dim intReturn As Integer = 0
		Try

			Dim fi As FileInfo
			Dim tmp As String = String.Empty
			' Loop through lstSource
			'For Each fi As FileInfo In m_lstSource
			For i As Integer = 0 To m_afi.Count - 1
				If m_afi(i) IsNot Nothing Then

					fi = m_afi(i) 'm_lstSource(i)

					' replace the source folder path with the dest path to get the full file string
					tmp = fi.FullName.Replace(m_strSourcePath, String.Empty)

					' Check if file exists at dest. If the file doesn't exist at dest, then copy the file
					'		OR if it does, check the last modified date. 
					'		if the last modified date on the source is greater (newer) than the dest
					'		then copy the file. 
					If Not File.Exists(m_strDestPath & tmp) OrElse
						fi.LastWriteTime.CompareTo(New FileInfo(m_strDestPath & tmp).LastWriteTime) > 0 Then
						If CopyFile(fi.FullName, m_strDestPath & tmp) Then
							Console.ForegroundColor = ConsoleColor.Green
							Console.WriteLine("Copied " & fi.FullName)
							Console.ForegroundColor = m_ccDefault
							tmp &= ": Success"
						Else
							Console.ForegroundColor = ConsoleColor.Red
							Console.WriteLine("Failed: " & fi.FullName)
							Console.ForegroundColor = m_ccDefault
							tmp &= ": Failed"
						End If
						m_lstCopied.Add(tmp)
						intReturn += 1

					End If

				End If
			Next
		Catch ex As Exception
			Console.WriteLine("Something when horrendously wrong, perhaps one Of your files is locked?" & vbCrLf & ex.Message)
		End Try
		Return intReturn
	End Function

	''' <summary>
	''' Checks if a file doesn't exists in the destination and copies or if it does exist, is the source older, then copy.
	''' </summary>
	''' <returns></returns>
	Public Function CopyOlder() As Integer
		Dim intReturn As Integer = 0
		Try

			Dim fi As FileInfo
			Dim tmp As String = String.Empty
			' Loop through lstSource
			'For Each fi As FileInfo In m_lstSource
			For i As Integer = 0 To m_afi.Count - 1
				If m_afi(i) IsNot Nothing Then

					fi = m_afi(i) 'm_lstSource(i)

					' replace the source folder path with the dest path to get the full file string
					tmp = fi.FullName.Replace(m_strSourcePath, String.Empty)
					' Check if file exists at dest. If the file doesn't exist at dest, then copy the file
					'		OR if it does, check the last modified date. 
					'		if the last modified date on the source is less (older) than the dest
					'		then copy the file. 
					If Not File.Exists(m_strDestPath & tmp) OrElse
						fi.LastWriteTime.CompareTo(New FileInfo(m_strDestPath & tmp).LastWriteTime) < 0 Then
						If CopyFile(fi.FullName, m_strDestPath & tmp) Then
							Console.ForegroundColor = ConsoleColor.Green
							Console.WriteLine("Copied " & fi.FullName)
							Console.ForegroundColor = m_ccDefault
							tmp &= ": Success"
						Else
							Console.ForegroundColor = ConsoleColor.Red
							Console.WriteLine("Failed: " & fi.FullName)
							Console.ForegroundColor = m_ccDefault
							tmp &= ": Failed"
						End If
						m_lstCopied.Add(tmp)
						intReturn += 1

					End If

				End If
			Next
		Catch ex As Exception
			Console.WriteLine("Something when horrendously wrong, perhaps one Of your files is locked?" & vbCrLf & ex.Message)
		End Try
		Return intReturn
	End Function

	''' <summary>
	''' Copies all files from the source to the destination
	''' </summary>
	''' <returns></returns>
	Public Function CopyAll() As Integer
		Dim intReturn As Integer = 0
		Try

			Dim fi As FileInfo
			Dim tmp As String = String.Empty
			' Loop through lstSource
			'For Each fi As FileInfo In m_lstSource
			For i As Integer = 0 To m_afi.Count - 1
				If m_afi(i) IsNot Nothing Then

					fi = m_afi(i) 'm_lstSource(i)

					' replace the source folder path with the dest path to get the full file string
					tmp = fi.FullName.Replace(m_strSourcePath, String.Empty)
					' Copy the file
					If CopyFile(fi.FullName, m_strDestPath & tmp) Then
						Console.ForegroundColor = ConsoleColor.Green
						Console.WriteLine("Copied " & fi.FullName)
						Console.ForegroundColor = m_ccDefault
						tmp &= ": Success"
					Else
						Console.ForegroundColor = ConsoleColor.Red
						Console.WriteLine("Failed: " & fi.FullName)
						Console.ForegroundColor = m_ccDefault
						tmp &= ": Failed"
					End If
					m_lstCopied.Add(tmp)
					intReturn += 1

				End If
			Next
		Catch ex As Exception
			Console.WriteLine("Something when horrendously wrong, perhaps one Of your files is locked?" & vbCrLf & ex.Message)
		End Try
		Return intReturn
	End Function

	''' <summary>
	''' Compares the file hashs, if they differ it copies, if the destination file doesn't exist, it copies
	''' </summary>
	''' <returns></returns>
	Public Function CopyHash() As Integer
		Dim intReturn As Integer = 0
		Try

			Dim fi As FileInfo
			Dim tmp As String = String.Empty
			' Loop through lstSource
			'For Each fi As FileInfo In m_lstSource
			For i As Integer = 0 To m_afi.Count - 1
				If m_afi(i) IsNot Nothing Then

					fi = m_afi(i) 'm_lstSource(i)

					' replace the source folder path with the dest path to get the full file string
					tmp = fi.FullName.Replace(m_strSourcePath, String.Empty)
					' Check if file exists at dest. If the file doesn't exist at dest, then copy the file
					'		OR if it does, check the last modified date. 
					'		if the last modified date on the source is greater (newer) than the dest
					'		then copy the file. 
					If Not File.Exists(m_strDestPath & tmp) OrElse
						Not New FileHash(fi).ToString.Equals(New FileHash(New FileInfo(m_strDestPath & tmp)).ToString) Then
						If CopyFile(fi.FullName, m_strDestPath & tmp) Then
							Console.ForegroundColor = ConsoleColor.Green
							Console.WriteLine("Copied " & fi.FullName)
							Console.ForegroundColor = m_ccDefault
							tmp &= ": Success"
						Else
							Console.ForegroundColor = ConsoleColor.Red
							Console.WriteLine("Failed: " & fi.FullName)
							Console.ForegroundColor = m_ccDefault
							tmp &= ": Failed"
						End If
						m_lstCopied.Add(tmp)
						intReturn += 1

					End If

				End If
			Next
		Catch ex As Exception
			Console.WriteLine("Something when horrendously wrong, perhaps one Of your files is locked?" & vbCrLf & ex.Message)
		End Try
		Return intReturn
	End Function

	''' <summary>
	''' Copies a file, creates all needed directories
	''' </summary>
	''' <param name="source"></param>
	''' <param name="dest"></param>
	''' <returns></returns>
	Private Function CopyFile(ByVal source As String, ByVal dest As String) As Boolean

		Dim blnReturn As Boolean = False

		Try
			Directory.CreateDirectory(New FileInfo(dest).DirectoryName)
			File.Copy(source, dest, True)
			blnReturn = True
		Catch fnf As FileNotFoundException
			Console.WriteLine("Source File not found: " & source)
		Catch ptl As PathTooLongException
			Console.WriteLine("The path is too long")
		Catch an As ArgumentException
			Console.WriteLine("Invalid source or destination file")
		Catch uax As UnauthorizedAccessException
			Console.WriteLine("Either the file is in use or you don't have access to modify it")
		Catch ex As Exception
			Console.WriteLine("Something really bad went wrong: " & ex.Message)
		End Try

		Return blnReturn

	End Function

	Private Function WriteLogFile(ByVal intFileCount As Integer, ByVal ts As TimeSpan) As Boolean

		Dim blnReturn As Boolean = False

		Try

			Dim astr(-1) As String
			astr.Add("File Copy Completed On " & DateTime.Now.ToShortDateString)
			If m_lstCopied.Count > 0 Then astr.AddRange(m_lstCopied.ToArray)
			astr.Add(intFileCount & " files copied")
			astr.Add("Time Elapsed: " & ts.Hours.ToString.PadLeft(2, "0") & ":" & ts.Minutes.ToString.PadLeft(2, "0") & ":" & ts.Seconds.ToString.PadLeft(2, "0") & "." & ts.Milliseconds.ToString.PadLeft(3, "0"))
			Dim strSave As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\CopyFileLogs\"
			Directory.CreateDirectory(strSave)
			Dim strfn As String

			strfn = "Copy_" & DateTime.Now.Year & DateTime.Now.Month.ToString.PadLeft(2, "0") & DateTime.Now.Day.ToString.PadLeft(2, "0") & "_" & DateTime.Now.Hour.ToString.PadLeft(2, "0") & DateTime.Now.Minute.ToString.PadLeft(2, "0") & DateTime.Now.Second.ToString.PadLeft(2, "0") & ".log"

			File.WriteAllLines(strSave & strfn, astr)

		Catch ex As Exception
			Console.WriteLine("Something went wrong. " & ex.Message)
		End Try

		Return blnReturn

	End Function


	''' <summary>
	''' Recursively loads all files in all subfolders into a given list object
	''' </summary>
	''' <param name="lst"></param>
	''' <param name="di"></param>
	Private Sub LoadFiles(ByRef lst As List(Of FileInfo), ByVal di As DirectoryInfo)

		Try
			Console.WriteLine("Loading Directory: " & di.FullName)

			lst.AddRange(di.GetFiles.ToList)

			For Each sd As DirectoryInfo In di.GetDirectories
				LoadFiles(lst, sd)
			Next
			GC.Collect()
		Catch ex As Exception
			Throw ex
		End Try

	End Sub

	''' <summary>
	''' Recursively loads all files in all subfolders into a given list object
	''' </summary>
	''' <param name="arr"></param>
	''' <param name="di"></param>
	Private Sub LoadFilesArray(ByRef arr() As FileInfo, ByVal di As DirectoryInfo)

		Try
			'Console.WriteLine("Loading Directory: " & di.FullName)

			arr.AddRange(di.GetFiles)
			'lst.AddRange(di.GetFiles.ToList)
			'For i As Integer = 0 To di.GetFiles.Count - 1
			'	arr.Add(di.GetFiles.GetValue(i))
			'Next

			For Each sd As DirectoryInfo In di.GetDirectories
				LoadFilesArray(arr, sd)
			Next
			GC.Collect()
		Catch ex As Exception
			Throw ex
		End Try

	End Sub


	Private Sub DisplayHelp()
		Dim intIndent As Integer = 3
		Console.WriteLine("Copy.exe [Copy From] [Copy To] -C [Compare Type<CopyNewer, CopyOlder, CopyHash, CopyAll(Default)>] --confirm -h")

		Console.WriteLine(vbCrLf)
		Console.WriteLine(Space(intIndent) & "About: This was written to make copying of newer files over odler ones easier then it spiraled out of control")
		Console.WriteLine(vbCrLf)
		Console.WriteLine("-C Compare Type can be values of: CopyNewer, CopyOlder, CopyAll")
		Console.WriteLine(Space(intIndent) & "CopyNewer: Copies and overwrites files that are newer in the source than the destination and file that don't exist")
		Console.WriteLine(Space(intIndent) & "CopyOlder: Copies and overwrites files that are older in the source than the destination and files that don't exist")
		Console.WriteLine(Space(intIndent) & "CopyHash: NOT IMPLEMENTED")
		Console.WriteLine(Space(intIndent) & "CopyAll: Copies and overwrites ALL files in the destination from the source")
		Console.WriteLine("--confirm check to make sure the file actually exists after copying")
		Console.WriteLine("-h Help (this menu)")
	End Sub



End Module

