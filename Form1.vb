Imports System.IO
Imports System.Text.RegularExpressions

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Check validity.
        If File.Exists(TextBox1.Text) = False Or Directory.Exists(TextBox2.Text) = False Then
            MsgBox("One of the paths are invalid, Please check.")
            Exit Sub
        End If

        'First of all get a list of all the files in that dir.
        Dim files As String() = TextBox1.Text.Split("|")

        'Super Globals.
        Dim mainPage As String = ""

        'Loop through all of them file by file now.
        For Each file As String In files
            'Get the files content.
            Dim fileCode As String = My.Computer.FileSystem.ReadAllText(file)

            'Create the file's markdown page.
            Dim filePage As String = ""

            'Add in the mainPage about this file.
            Dim tmpMain As String = ""
            tmpMain += ("## [FILE_NAME](FILE_NAME)" + vbCrLf + vbCrLf).Replace("FILE_NAME", Path.GetFileNameWithoutExtension(file))

            'Loop through all the pawndoc's
            For Each mtch As Match In Regex.Matches(fileCode, "\/\*\*([\s\S]*?)\*\*\/")
                'Parse the doc.
                Dim pwnDoc As New PawnDoc(mtch.Groups(1).Value)

                'Create the var to store the template.
                Dim funcMark As String = My.Resources.funcTemplate1

                'Check if valid.
                If pwnDoc.Summary = "" Or pwnDoc.Summary = "-" Then Continue For

                'Replace all variables.
                funcMark = funcMark.Replace("FUNCTION_NAME", prepareFuncName(pwnDoc.Summary))
                funcMark = funcMark.Replace("TEXT_IN_REMARKS_SECTION", prepareString(pwnDoc.Remarks, False))
                funcMark = funcMark.Replace("LINES_OF_TEXT_IN_RETURNS_SECTION", prepareString(pwnDoc.Returns))
                Dim pars As String = ""
                For Each par In pwnDoc.Parameters
                    pars += (">" + vbTab + "* `PARAM_NAME`: PARAM_INFO" + vbCrLf).Replace("PARAM_NAME", par.paramName).Replace("PARAM_INFO", par.paramExplain)
                Next
                funcMark = funcMark.Replace("LINES_OF_TEXT_IN_PARAMS_SECTION", pars.Trim)

                'Append to the file's page.
                filePage += funcMark

                'Append to the main page.
                tmpMain += ("* [FUNCTION_NAME_UP](FILE_NAME#FUNCTION_NAME_UNDER)" + vbCrLf).Replace("FILE_NAME", Path.GetFileNameWithoutExtension(file)).Replace("FUNCTION_NAME_UP", prepareFuncName(pwnDoc.Summary)).Replace("FUNCTION_NAME_UNDER", prepareFuncName(pwnDoc.Summary, True))
            Next

            'Now after we have all the files functions save it.
            If filePage <> "" Then
                'Save the actual file.
                My.Computer.FileSystem.WriteAllText(TextBox2.Text + "/" + Path.GetFileNameWithoutExtension(file) + ".md", filePage, False)

                'Append to main file.
                mainPage += tmpMain + vbCrLf + vbCrLf
            End If
        Next

        'Now save the main file.
        My.Computer.FileSystem.WriteAllText(TextBox2.Text + "/home.md", mainPage, False)
        MsgBox("Done!")
    End Sub

    Function prepareFuncName(ByVal name As String, Optional ByVal isLower As Boolean = False) As String
        prepareFuncName = ""

        prepareFuncName = Regex.Replace(name, "\(.*\);?", "").Replace(vbCrLf, "|")

        If prepareFuncName.Contains(":") Then
            prepareFuncName = prepareFuncName.Remove(0, name.IndexOf(":") + 1)
        End If

        If isLower = True Then
            prepareFuncName = prepareFuncName.ToLower()
        End If

        Return prepareFuncName
    End Function

    Private Function prepareString(str As String, Optional asterisk As Boolean = True) As String
        prepareString = ""

        If str Is Nothing Then Return ""

        Dim lines As String() = str.Split(vbCrLf)

        For Each line In lines
            If asterisk Then
                prepareString += ">" + vbTab + "* " + line.Trim + vbCrLf
            Else
                prepareString += ">" + vbTab + line.Trim + vbCrLf
            End If
        Next

        Return prepareString.Trim
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            TextBox1.Text = ""
            For Each Path As String In OpenFileDialog1.FileNames
                TextBox1.Text += Path + "|"
            Next

            TextBox1.Text = TextBox1.Text.Remove(TextBox1.Text.Length - 1, 1)
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            TextBox2.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub
End Class
