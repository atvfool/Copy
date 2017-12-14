Imports System.Runtime.CompilerServices

Public Module Extensions
	<Extension()>
	Public Sub Add(Of T)(ByRef arr As T(), item As T)
		If arr IsNot Nothing Then
			Array.Resize(arr, arr.Length + 1)
			arr(arr.Length - 1) = item
		Else
			ReDim arr(0)
			arr(0) = item
		End If
	End Sub

	<Extension()>
	Public Sub AddRange(Of T)(ByRef arr As T(), item As T())

		If arr IsNot Nothing Then
			'Dim intStartIndex As Integer = arr.Length
			'Array.Resize(arr, arr.Length + item.Length)
			'arr(arr.Length - 1) =
			For i As Integer = 0 To item.Length - 1
				arr.Add(item(i))
			Next
		Else
			' Arr doesn't exist yet so just a straight copy
			ReDim arr(item.Length)
			For I As Integer = 0 To item.Length - 1
				arr(I) = item(I)
			Next
		End If

	End Sub

	<Extension()>
	Public Sub Print(Of T)(ByRef arr As T())
		For i As Integer = 0 To arr.Length - 1

			If arr(i) IsNot Nothing Then
				Console.WriteLine(arr(i))
			End If

		Next
	End Sub

End Module
