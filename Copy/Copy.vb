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
	Private m_blnIsSourceAFile As Boolean = False
	Private m_blnIsDestFilename As Boolean = False
	Private m_strFilter As String = String.Empty
	Private m_ccDefault As ConsoleColor
	Private m_CompareType As New CompareType
	Private Enum CompareType
		CopyNewer = 0
		CopyOlder
		CopyDiffHash
		CopyAll
	End Enum

	''' <summary>
	''' The Starting point of the application.
	''' </summary>
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
		If args.Contains("-C") OrElse args.Contains("-c") Then
			For i As Integer = 0 To args.Length - 1
				If args(i).ToUpper = "-C" Then
					m_CompareType = DirectCast([Enum].Parse(GetType(CompareType), args(i + 1)), CompareType)
					Exit For
				End If
			Next
		Else
			m_CompareType = CompareType.CopyAll
		End If

		' Get the filter
		If args.Contains("-F") OrElse args.Contains("-f") Then
			For i As Integer = 0 To args.Length - 1
				If args(i).ToUpper = "-F" Then
					m_strFilter = args(i + 1)
					Exit For
				End If
			Next
		End If

		m_strSourcePath = args(1)
		m_strDestPath = args(2)

		' Check if the path leading up to the source file/directory exists
		If Directory.Exists(Path.GetDirectoryName(m_strSourcePath)) Then
			' Check if it's a file or directory
			Dim fa As FileAttributes = File.GetAttributes(m_strSourcePath)
			If Not File.GetAttributes(m_strSourcePath).HasFlag(FileAttributes.Directory) Then
				m_blnIsSourceAFile = True
			End If
		Else
			Console.WriteLine("Path doesn't exist")
			End
		End If

		' Check if dest path ends in extension, if so then use this as the file name, otherwise use the source file name
		If Path.GetExtension(m_strDestPath) <> String.Empty Then
			m_blnIsDestFilename = True
		End If

		' Check if source is not a file and dest is a filename then we have a problem
		If Not m_blnIsSourceAFile And m_blnIsDestFilename Then
			Console.WriteLine("You can't put a directory into a file... yet")
			End
		End If





		Dim intCount As Integer = 0
		Dim dtStart As DateTime = DateTime.Now
		Dim dtFinish As DateTime
		If (m_blnIsSourceAFile AndAlso IO.File.Exists(m_strSourcePath)) _
			OrElse IO.Directory.Exists(m_strSourcePath) Then

			LoadFilesArray(m_afi, m_strSourcePath, m_strFilter)
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
	''' <returns>Number of files found</returns>
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
					tmp = If(m_blnIsDestFilename, String.Empty, fi.FullName.Replace(Path.GetDirectoryName(m_strSourcePath), String.Empty))

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
	''' <returns>Number of files found</returns>
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
					tmp = If(m_blnIsDestFilename, String.Empty, fi.FullName.Replace(Path.GetDirectoryName(m_strSourcePath), String.Empty))
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
	''' <returns>Number of files found</returns>
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
					tmp = If(m_blnIsDestFilename, String.Empty, fi.FullName.Replace(Path.GetDirectoryName(m_strSourcePath), String.Empty))
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
	''' <returns>Number of files found</returns>
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
					tmp = If(m_blnIsDestFilename, String.Empty, fi.FullName.Replace(Path.GetDirectoryName(m_strSourcePath), String.Empty))
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
	''' <returns>Number of files found</returns>
	Private Function CopyFile(ByVal source As String, ByVal dest As String) As Boolean

		Dim blnReturn As Boolean = False

		Try
			Dim fi As New FileInfo(dest)
			If fi.Exists Then fi.IsReadOnly = False
			Directory.CreateDirectory(fi.DirectoryName)

			File.Copy(source, fi.FullName, True)
			blnReturn = True
		Catch fnf As FileNotFoundException
			WriteErrorLog("Source File not found: " & source)
		Catch ptl As PathTooLongException
			WriteErrorLog("The path is too long" & vbCrLf & "Source:" & source & vbCrLf & "Destination: " & dest)
		Catch an As ArgumentException
			WriteErrorLog("Invalid source or destination file")
		Catch uax As UnauthorizedAccessException
			WriteErrorLog("Either the file is in use or you don't have access to modify it")
		Catch ex As Exception
			WriteErrorLog("Something really bad went wrong: " & ex.Message)
		End Try

		Return blnReturn

	End Function

	''' <summary>
	''' Writes a log file with file count, time elapsed and current date
	''' </summary>
	''' <param name="intFileCount">Number of files affected</param>
	''' <param name="ts">Elapsed time</param>
	''' <returns>True if successful</returns>
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

			blnReturn = True

		Catch ex As Exception
			Console.WriteLine("Something went wrong. " & ex.Message)
		End Try

		Return blnReturn

	End Function

	''' <summary>
	''' Writes an error message to a log file
	''' </summary>
	''' <param name="strMessage">The error message</param>
	''' <param name="blnPrintMessage">True to print in the console window</param>
	''' <returns>True if file write is a success</returns>
	Private Function WriteErrorLog(ByVal strMessage As String, Optional ByVal blnPrintMessage As Boolean = True) As Boolean

		Dim blnResult As Boolean = False

		Try
			Dim astr(-1) As String
			astr.Add("An error occured at " & Date.Now & ". The error message is as follows:")
			astr.Add(strMessage)
			Dim strSave As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\CopyFileLogs\"
			Directory.CreateDirectory(strSave)
			Dim strfn As String

			strfn = "Copy_ErrorLog_" & DateTime.Now.Year & DateTime.Now.Month.ToString.PadLeft(2, "0") & DateTime.Now.Day.ToString.PadLeft(2, "0") & "_" & DateTime.Now.Hour.ToString.PadLeft(2, "0") & DateTime.Now.Minute.ToString.PadLeft(2, "0") & DateTime.Now.Second.ToString.PadLeft(2, "0") & ".log"

			File.WriteAllLines(strSave & strfn, astr)

			If blnPrintMessage Then
				Console.WriteLine(strMessage)
			End If

			blnResult = True

		Catch ex As Exception
			Console.WriteLine("There was an error writing the Log File: " & ex.Message)
		End Try

		Return blnResult

	End Function


	''' <summary>
	''' Recursively loads all files in all subfolders into a given list object
	''' </summary>
	''' <param name="lst">List of file info</param>
	''' <param name="di">Current Direcotry info</param>
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
	''' Recursively loads all files in all subfolders into a given list object unless the source is a single file
	''' </summary>
	''' <param name="arr">The array to load the files into</param>
	''' <param name="strSource">Where to load the files from</param>
	Private Sub LoadFilesArray(ByRef arr() As FileInfo, ByVal strSource As String, ByVal strFilter As String)

		Try
			'Console.WriteLine("Loading Directory: " & di.FullName)

			If m_blnIsSourceAFile Then
				arr.Add(New FileInfo(strSource))
			Else
				Dim di As New DirectoryInfo(strSource)
				If strFilter.Trim.Equals(String.Empty) Then
					arr.AddRange(di.GetFiles())
				Else
					arr.AddRange(di.GetFiles(strFilter))
				End If

				'lst.AddRange(di.GetFiles.ToList)
				'For i As Integer = 0 To di.GetFiles.Count - 1
				'	arr.Add(di.GetFiles.GetValue(i))
				'Next

				For Each sd As DirectoryInfo In di.GetDirectories
					LoadFilesArray(arr, sd.FullName, strFilter)
				Next
			End If
			GC.Collect()
		Catch ex As Exception
			Throw ex
		End Try

	End Sub

	''' <summary>
	''' Displays a help message in the console
	''' </summary>
	Private Sub DisplayHelp()
		Dim intIndent As Integer = 3
		Console.WriteLine("Copy.exe [Path/File From] [Path/File To] -C [Compare Type<CopyNewer, CopyOlder, CopyHash, CopyAll(Default)>] -F ""<String for filename and extension(.abc)>"" -h")

		Console.WriteLine(vbCrLf)
		Console.WriteLine(Space(intIndent) & "About: This was written to make copying of newer files over odler ones easier then it spiraled out of control")
		Console.WriteLine(vbCrLf)
		Console.WriteLine("-C Compare Type can be values of: CopyNewer, CopyOlder, CopyAll")
		Console.WriteLine(Space(intIndent) & "CopyNewer: Copies and overwrites files that are newer in the source than the destination and file that don't exist")
		Console.WriteLine(Space(intIndent) & "CopyOlder: Copies and overwrites files that are older in the source than the destination and files that don't exist")
		Console.WriteLine(Space(intIndent) & "CopyHash: Overwrites destination files based on whether their hash's match")
		Console.WriteLine(Space(intIndent) & "CopyAll: Copies and overwrites ALL files in the destination from the source")
		Console.WriteLine(vbCrLf)
		Console.WriteLine("-F ""<String for filename and extension(.abc)>""")
		Console.WriteLine(Space(intIndent) & "Extensions must start with a period e.g. for Word Document ""*.doc"" or ""*.docx""")
		Console.WriteLine(Space(intIndent) & "You can only filter when copying directories")
		'Console.WriteLine("--confirm check to make sure the file actually exists after copying")
		Console.WriteLine("-h Help (this menu)")
	End Sub



End Module

