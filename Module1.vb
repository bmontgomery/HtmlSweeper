Imports HtmlAgilityPack

Module Module1

	Private mProcessDirectory As String
	Private mFileMatchPattern As String

	Sub Main()
		Try

			ParseCommandLineArguments()

			If Not String.IsNullOrEmpty(mProcessDirectory) Then
				RecurseDirectories(mProcessDirectory)
			End If

		Catch ex As Exception
			Console.Write(ex.Message + Environment.NewLine + ex.StackTrace)
		End Try

		Console.ReadLine()

	End Sub

	Private Sub ParseCommandLineArguments()

		If Environment.GetCommandLineArgs().Length > 1 Then

			mProcessDirectory = Environment.GetCommandLineArgs(1)
			mFileMatchPattern = Environment.GetCommandLineArgs(2)

		End If

	End Sub

	Private Sub RecurseDirectories(ByVal directoryPath As String)

		For Each directory As String In IO.Directory.GetDirectories(directoryPath)
			RecurseDirectories(directory)
		Next

		For Each file As String In IO.Directory.GetFiles(directoryPath)
			If Text.RegularExpressions.Regex.IsMatch(file, mFileMatchPattern) Then
				FixHtml(file)
			End If
		Next

	End Sub

	Private Sub FixHtml(ByVal fileName As String)

		Dim doc As New HtmlDocument()
		doc.Load(fileName)
		Dim culpritNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//td/span[@class='frmFldLbl']")

		If culpritNodes IsNot Nothing Then
			For Each culpritNode As HtmlNode In culpritNodes

				Dim culpritNodeIndex As Int32 = culpritNode.ParentNode.ChildNodes.IndexOf(culpritNode)
				Dim culpritNodeText As String = culpritNode.InnerHtml
				Dim parentTdClassAtt As HtmlAttribute = culpritNode.ParentNode.Attributes("class")

				If Not parentTdClassAtt.Value.Contains("frmFldLbl") Then

					If Not String.IsNullOrEmpty(parentTdClassAtt.Value) Then parentTdClassAtt.Value += " "
					parentTdClassAtt.Value += "frmFldLbl"

				End If

				'Dim replacementNode As New HtmlNode(HtmlNodeType.Text, doc, 0)
				'replacementNode.InnerHtml = culpritNodeText
				'culpritNode.ParentNode.ChildNodes.Insert(culpritNodeIndex, replacementNode)
				culpritNode.Remove()

			Next
		End If

		doc.Save(fileName)

	End Sub

End Module
