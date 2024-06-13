Imports System.ComponentModel

Public Class Form1
    Private Const newLineCharacter As Char = vbLf
    Private Const RequestSize As UInt64 = 5000
    Private Const RequestRate As UInt64 = 950
    Private ReadOnly httpClient As New Net.Http.HttpClient()
    Private voices As SortedDictionary(Of String, SortedDictionary(Of String, SortedDictionary(Of String, SsmlGender)))
    Private suspended1 As Boolean = False
    Private suspended2 As Boolean = False
    Private statCalculationThread As New Threading.Thread(Sub()
                                                          End Sub)
    Private requestCount As UInt64 = 0
    Private lastRequestTime As Date = Date.MinValue
    Private processes As UInt64 = 0
    Private Async Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Width = Panel4.ClientSize.Width - 1
        ComboBox1.Top = Panel4.ClientSize.Height / 2 - ComboBox1.Height / 2
        ComboBox1.Anchor = AnchorStyles.Left Or AnchorStyles.Right
        ComboBox2.Width = Panel5.ClientSize.Width - 1
        ComboBox2.Top = Panel5.ClientSize.Height / 2 - ComboBox2.Height / 2
        ComboBox2.Anchor = AnchorStyles.Left Or AnchorStyles.Right
        TextBox1.Left = 3
        TextBox1.Width = Panel1.ClientSize.Width - 6
        TextBox1.Top = Panel1.ClientSize.Height / 2 - TextBox1.Height / 2
        TextBox1.Anchor = AnchorStyles.Left Or AnchorStyles.Right
        TextBox2.Left = 3
        TextBox2.Width = Panel2.ClientSize.Width - 6
        TextBox2.Top = Panel2.ClientSize.Height / 2 - TextBox2.Height / 2
        TextBox2.Anchor = AnchorStyles.Left Or AnchorStyles.Right
        ComboBox3.Width = Panel6.ClientSize.Width - 1
        ComboBox3.Top = Panel6.ClientSize.Height / 2 - ComboBox3.Height / 2
        ComboBox3.Anchor = AnchorStyles.Left Or AnchorStyles.Right
        WindowState = FormWindowState.Maximized
        ComboBox3.SelectedIndex = 1
        voices = Await GetVoices()
        ComboBox1.Items.AddRange(voices.Keys.ToArray())
        ComboBox1.SelectedIndex = 0
        Reset0()
        ComboBox2.SelectedIndex = 0
        Reset1()
        ComboBox4.SelectedIndex = 0
        ComboBox1.SelectedItem = "Wavenet"
        ComboBox2.SelectedItem = "en-US"
        ComboBox4.SelectedItem = "en-US-Wavenet-J (Male)"
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        ComboBox4.Enabled = True
        Button2.Enabled = True
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.ShowDialog()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox2.Text = "0" Then
            Dim thread As New Threading.Thread(Sub()
                                                   Const WaitTime As UInt64 = 150
                                                   For i As UInt64 = 0 To 5
                                                       Me.Invoke(Sub()
                                                                     TextBox2.BackColor = Color.Red
                                                                     Panel2.BackColor = Color.Red
                                                                 End Sub)
                                                       Threading.Thread.Sleep(WaitTime)
                                                       Me.Invoke(Sub()
                                                                     TextBox2.BackColor = Color.White
                                                                     Panel2.BackColor = Color.White
                                                                 End Sub)
                                                       Threading.Thread.Sleep(WaitTime)
                                                   Next
                                                   Me.Invoke(Sub()
                                                                 Me.Refresh()
                                                             End Sub)
                                               End Sub)
            thread.Start()
        Else
            Dim fileName As String = SaveFileDialog1.FileName
            fileName = fileName.Split("\")(fileName.Split("\").Length - 1)
            fileName = fileName.Substring(0, fileName.Length - fileName.Split(".")(fileName.Split(".").Length - 1).Length - 1)
            SaveFileDialog1.FileName = fileName
            SaveFileDialog1.ShowDialog()
        End If
    End Sub
    Private Async Function GetAudio(Text As String, Voice As String, Pitch As Integer, Speed As Double) As Task(Of Byte())
        Dim request As String = "{""audioConfig"":{""audioEncoding"":""MP3"",""pitch"":" & Pitch & ",""speakingRate"":" & Speed & "},""input"":{""text"":" & Newtonsoft.Json.JsonConvert.SerializeObject(Text) & "},""voice"":{""languageCode"":""en-US"",""name"":""" & Voice & """}}"
        Dim httpContent As New Net.Http.StringContent(request)
        Dim rawResponse As Net.Http.HttpResponseMessage = Await httpClient.PostAsync("https://texttospeech.googleapis.com/v1beta1/text:synthesize?key=AIzaSyAOsImUr0W7AjLx9dn4fbI3IL7q6M0dbFE", httpContent)
        Dim response As String = Await rawResponse.Content.ReadAsStringAsync()
        Dim result As String = Newtonsoft.Json.JsonConvert.DeserializeObject(Of ResponseType0)(response).audioContent
        Return Convert.FromBase64String(result)
    End Function
    Private Async Function GetRawVoices() As Task(Of RawVoiceInformation())
        Dim rawResponse As Net.Http.HttpResponseMessage = Await httpClient.GetAsync("https://texttospeech.googleapis.com/v1/voices?key=AIzaSyAOsImUr0W7AjLx9dn4fbI3IL7q6M0dbFE")
        Dim response As String = Await rawResponse.Content.ReadAsStringAsync()
        Return Newtonsoft.Json.JsonConvert.DeserializeObject(Of ResponseType1)(response).voices
    End Function
    Private Async Function GetVoices() As Task(Of SortedDictionary(Of String, SortedDictionary(Of String, SortedDictionary(Of String, SsmlGender))))
        Dim voices As New SortedDictionary(Of String, SortedDictionary(Of String, SortedDictionary(Of String, SsmlGender)))()
        Dim rawVoices As RawVoiceInformation() = Await GetRawVoices()
        For Each rawVoice As RawVoiceInformation In rawVoices
            Dim languageCode As String = rawVoice.languageCodes(0)
            Dim nameSplit As String() = rawVoice.name.Split("-"c)
            Dim network As String = nameSplit(nameSplit.Length - 2)
            Dim gender As SsmlGender = If(rawVoice.ssmlGender = "MALE", SsmlGender.male, SsmlGender.female)
            If Not voices.ContainsKey(network) Then voices.Add(network, New SortedDictionary(Of String, SortedDictionary(Of String, SsmlGender))())
            If Not voices(network).ContainsKey(languageCode) Then voices(network).Add(languageCode, New SortedDictionary(Of String, SsmlGender)())
            voices(network)(languageCode)(rawVoice.name) = gender
        Next
        Return voices
    End Function
    Private Function GetLength(Text As String, Voice As String, Pitch As Integer, Speed As Double) As Boolean
        Return RequestSize >= System.Text.Encoding.UTF8.GetBytes(("{""audioConfig"":{""audioEncoding"":""MP3"",""pitch"":" & Pitch & ",""speakingRate"":" & Speed & "},""input"":{""text"":" & Newtonsoft.Json.JsonConvert.SerializeObject(Text) & "},""voice"":{""languageCode"":""en-US"",""name"":""" & Voice & """}}")).Length
    End Function
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Reset0()
        Changed()
    End Sub
    Private Sub Reset0()
        Dim selectedItem As Object = ComboBox2.SelectedItem
        If ComboBox1.SelectedIndex <> -1 Then
            ComboBox2.Items.Clear()
            ComboBox2.Items.AddRange(voices(ComboBox1.SelectedItem).Keys.ToArray())
            If IsNothing(selectedItem) Then Return
            If ComboBox2.Items.Contains(selectedItem) Then
                ComboBox2.SelectedItem = selectedItem
            Else
                ComboBox2.SelectedIndex = ComboBox2.Items.Count - 1
            End If
        End If
    End Sub
    Private Sub Reset1()
        Dim selectedItem As Object = ComboBox4.SelectedItem
        If ComboBox2.SelectedIndex <> -1 Then
            ComboBox4.Items.Clear()
            For Each i As KeyValuePair(Of String, SsmlGender) In voices(ComboBox1.SelectedItem)(ComboBox2.SelectedItem)
                ComboBox4.Items.Add(i.Key & " (" & If(i.Value = 1, "Male", "Female") & ")")
            Next
            If IsNothing(selectedItem) Then Return
            If ComboBox4.Items.Contains(selectedItem) Then
                ComboBox4.SelectedItem = selectedItem
            Else
                ComboBox4.SelectedIndex = ComboBox4.Items.Count - 1
            End If
        End If
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If Not suspended1 Then
            Dim currentPosition As Integer = TextBox1.SelectionStart
            Dim text0 As String = TextBox1.Text
            Dim permittedCharacters As Char() = {"0"c, "1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c}
            Dim text1 As String = ""
            Dim sign As Boolean = True
            For Each character As Char In text0
                If permittedCharacters.Contains(character) Then
                    text1 &= character
                Else
                    If text1.Length <= currentPosition Then
                        currentPosition -= 1
                    End If
                    If character = "-" Then
                        sign = Not sign
                    ElseIf character = "+" Then
                        sign = True
                    End If
                End If
            Next
            If text1 = "" Then
                text1 = "0"
            ElseIf Integer.Parse(text1) > 20 Then
                text1 = "20"
            Else
                While text1.StartsWith("0")
                    text1 = text1.Substring(1)
                End While
                If text1 = "" Then text1 = "0"
            End If
            If Not sign Then
                text1 = "-" & text1
                currentPosition += 1
            End If
            suspended1 = True
            TextBox1.Text = text1
            suspended1 = False
            If text1 = "0" Then
                TextBox1.SelectionStart = 1
            ElseIf text1 = "-0" Then
                TextBox1.SelectionStart = 2
            Else
                TextBox1.SelectionStart = currentPosition
            End If
        End If
    End Sub
    Private Sub TextBox1_LostFocus(sender As Object, e As EventArgs) Handles TextBox1.LostFocus
        If TextBox1.Text = "-0" Then TextBox1.Text = "0"
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If Not suspended2 Then
            Dim currentPosition As Integer = TextBox2.SelectionStart
            Dim text0 As String = TextBox2.Text
            Dim permittedCharacters As Char() = {"0"c, "1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c}
            Dim text1 As String = ""
            Dim seenDecimal As Boolean = False
            For Each character As Char In text0
                If permittedCharacters.Contains(character) Then
                    text1 &= character
                ElseIf character = "."c And (Not seenDecimal) Then
                    text1 &= "."
                    seenDecimal = True
                Else
                    If text1.Length <= currentPosition Then
                        currentPosition -= 1
                    End If
                End If
            Next
            If text1 = "" Then
                text1 = "0"
            ElseIf text1.StartsWith(".") Then
                text1 = "0" & text1
            ElseIf Double.Parse(text1) > 10 Then
                text1 = "10"
            Else
                While text1.StartsWith("0") And If(text1.Contains("."c), text1.IndexOf("."c) > 1, True)
                    text1 = text1.Substring(1)
                End While
                If text1 = "" Then text1 = "0"
            End If
            suspended2 = True
            TextBox2.Text = text1
            suspended2 = False
            If text1 = "0" Then
                TextBox2.SelectionStart = 1
            Else
                TextBox2.SelectionStart = currentPosition
            End If
        End If
    End Sub
    Private Sub TextBox2_LostFocus(sender As Object, e As EventArgs) Handles TextBox2.LostFocus
        If TextBox2.Text.EndsWith("."c) Then TextBox2.Text = TextBox2.Text.Substring(0, TextBox2.Text.Length - 1)
    End Sub
    Private Sub Changed()
        statCalculationThread.Abort()
        Try
            statCalculationThread.Join()
        Catch ex As Threading.ThreadStateException
        End Try
        Dim text As String = RichTextBox1.Text
        Dim compression As CompressionLevel = ComboBox3.SelectedIndex
        Dim network As String = ComboBox1.SelectedItem
        Label5.Text = "Characters: " & text.Length
        If (compression = CompressionLevel.Level1 And text.Length > 15000000) Or (compression = CompressionLevel.Level2 And text.Length > 50000) Then
            Label4.Text = "Words: ..."
            Label6.Text = "Estimated Cost: ..."
            Label7.Text = "Consigning Characters: ..."
        End If
        statCalculationThread = New Threading.Thread(Sub()
                                                         Dim length As Integer = text.Length
                                                         Dim wordText As String = text.Trim()
                                                         While wordText.Contains("  ")
                                                             wordText = wordText.Replace("  ", " ")
                                                         End While
                                                         Dim words As UInt64
                                                         For Each character As Char In wordText
                                                             If character = " "c Then
                                                                 words += 1
                                                             End If
                                                         Next
                                                         text = Compress(text, compression)
                                                         Dim estimatedCost As Double
                                                         Select Case network
                                                             Case "Standard"
                                                                 estimatedCost = text.Length * 0.000004
                                                             Case "Wavenet"
                                                                 estimatedCost = text.Length * 0.000016
                                                             Case "Neural2"
                                                                 estimatedCost = text.Length * 0.000016
                                                             Case Else
                                                                 estimatedCost = text.Length * 0.000016
                                                         End Select
                                                         Dim stringWords As String = "Words: " & If(wordText.Length = 0, "0", words + 1)
                                                         Dim stringEstimatedCost As String = "Estimated Cost: $" & estimatedCost.ToString("F5")
                                                         Dim stringConsigningCharacters As String = "Consigning Characters: " & text.Length
                                                         Me.Invoke(Sub()
                                                                       Label4.Text = stringWords
                                                                       Label6.Text = stringEstimatedCost
                                                                       Label7.Text = stringConsigningCharacters
                                                                   End Sub)
                                                     End Sub)
        statCalculationThread.Start()
    End Sub
    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged
        RichTextBox1.BackColor = Color.White
        RichTextBox1.ForeColor = Color.Black
        RichTextBox1.Font = New Font("Microsoft Sans Serif", 8)
        Changed()
    End Sub
    Private Function Compress(value As String, Compression As CompressionLevel) As String
        If Compression = CompressionLevel.Level0 Then Return value
        Dim alphabet As Char() = {"a"c, "b"c, "c"c, "d"c, "e"c, "f"c, "g"c, "h"c, "i"c, "j"c, "k"c, "l"c, "m"c, "n"c, "o"c, "p"c, "q"c, "r"c, "s"c, "t"c, "u"c, "v"c, "w"c, "x"c, "y"c, "z"c}
        Dim numbers As Char() = {"0"c, "1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c}
        Dim newValue As String = ""
        If Compression = CompressionLevel.Level2 Then
            value = value.ToLower()
            If value.Length >= 2 Then
                newValue = value(0)
                For i As UInt64 = 1 To value.Length - 1
                    If alphabet.Contains(value(i - 1)) Then
                        newValue &= value(i)
                    Else
                        newValue &= ToUpper(value(i))
                    End If
                Next
            Else
                newValue = value
            End If
            value = newValue
            newValue = ""
        End If
        value = value.Replace(newLineCharacter, " ")
        value = value.Trim()
        While value.Contains("  ")
            value = value.Replace("  ", " ")
        End While
        If Compression = CompressionLevel.Level2 Then
            For i As Int64 = 0 To value.Length - 1
                If value(i) = " "c Then
                    If value(i - 1) = ToUpper(value(i - 1)) Or ((numbers.Contains(value(i - 1)) Or value(i - 1) = "."c) And numbers.Contains(value(i + 1))) Then
                        newValue &= value(i)
                    End If
                Else
                    newValue &= value(i)
                End If
            Next
            Return newValue
        Else
            Return value
        End If
    End Function
    Public alphabet As Char() = {"a"c, "b"c, "c"c, "d"c, "e"c, "f"c, "g"c, "h"c, "i"c, "j"c, "k"c, "l"c, "m"c, "n"c, "o"c, "p"c, "q"c, "r"c, "s"c, "t"c, "u"c, "v"c, "w"c, "x"c, "y"c, "z"c}
    Public numbers As Char() = {"0"c, "1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c}
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Reset1()
        Changed()
    End Sub
    Public Function ToLower(value As Char) As Char
        Select Case value
            Case "A"c
                Return "a"c
            Case "B"c
                Return "b"c
            Case "C"c
                Return "c"c
            Case "D"c
                Return "d"c
            Case "E"c
                Return "e"c
            Case "F"c
                Return "f"c
            Case "G"c
                Return "g"c
            Case "H"c
                Return "h"c
            Case "I"c
                Return "i"c
            Case "J"c
                Return "j"c
            Case "K"c
                Return "k"c
            Case "L"c
                Return "l"c
            Case "M"c
                Return "m"c
            Case "N"c
                Return "n"c
            Case "O"c
                Return "o"c
            Case "P"c
                Return "p"c
            Case "Q"c
                Return "q"c
            Case "R"c
                Return "r"c
            Case "S"c
                Return "s"c
            Case "T"c
                Return "t"c
            Case "U"c
                Return "u"c
            Case "V"c
                Return "v"c
            Case "W"c
                Return "w"c
            Case "X"c
                Return "x"c
            Case "Y"c
                Return "y"c
            Case "Z"c
                Return "z"c
            Case Else
                Return value
        End Select
    End Function
    Public Function ToUpper(value As Char) As Char
        Select Case value
            Case "a"c
                Return "A"c
            Case "b"c
                Return "B"c
            Case "c"c
                Return "C"c
            Case "d"c
                Return "D"c
            Case "e"c
                Return "E"c
            Case "f"c
                Return "F"c
            Case "g"c
                Return "G"c
            Case "h"c
                Return "H"c
            Case "i"c
                Return "I"c
            Case "j"c
                Return "J"c
            Case "k"c
                Return "K"c
            Case "l"c
                Return "L"c
            Case "m"c
                Return "M"c
            Case "n"c
                Return "N"c
            Case "o"c
                Return "O"c
            Case "p"c
                Return "P"c
            Case "q"c
                Return "Q"c
            Case "r"c
                Return "R"c
            Case "s"c
                Return "S"c
            Case "t"c
                Return "T"c
            Case "u"c
                Return "U"c
            Case "v"c
                Return "V"c
            Case "w"c
                Return "W"c
            Case "x"c
                Return "X"c
            Case "y"c
                Return "Y"c
            Case "z"c
                Return "Z"c
            Case Else
                Return value
        End Select
    End Function
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Changed()
    End Sub
    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Refresh()
    End Sub
    Private Sub Form1_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        Me.Refresh()
    End Sub
    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Dim found As Boolean = True
        If e.Control Then
            Select Case e.KeyCode
                Case Keys.L
                    Button1.PerformClick()
                Case Keys.O
                    Button1.PerformClick()
                Case Keys.S
                    Button2.PerformClick()
                Case Else
                    found = False
            End Select
        ElseIf e.Alt Then
            Select Case e.KeyCode
                Case Keys.N
                    ComboBox1.Focus()
                Case Keys.V
                    ComboBox2.Focus()
                Case Keys.P
                    TextBox1.Focus()
                Case Keys.S
                    TextBox2.Focus()
                Case Keys.C
                    ComboBox3.Focus()
                Case Keys.T
                    RichTextBox1.Focus()
                Case Keys.I
                    RichTextBox1.Focus()
                Case Else
                    found = False
            End Select
        Else
            found = False
        End If
        e.Handled = found
        e.SuppressKeyPress = found
    End Sub
    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        Dim newFileName As String = OpenFileDialog1.FileName
        RichTextBox1.Text = IO.File.ReadAllText(newFileName)
        SaveFileDialog1.FileName = newFileName
        SaveFileDialog1.FileName = newFileName.Split("\")(newFileName.Split("\").Length - 1)
    End Sub
    Private Sub SaveFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk
        Dim newThread As New Threading.Thread(Async Sub()
                                                  Me.Invoke(Sub()
                                                                processes += 1
                                                            End Sub)
                                                  Dim strings As New List(Of String)()
                                                  Dim destination As String = Me.Invoke(Function()
                                                                                            Return SaveFileDialog1.FileName
                                                                                        End Function)
                                                  Dim text As String = Me.Invoke(Function()
                                                                                     Return RichTextBox1.Text
                                                                                 End Function)
                                                  Dim Voice As String = Me.Invoke(Function()
                                                                                      Return ComboBox4.SelectedItem.ToString().Split(" "c)(0)
                                                                                  End Function).ToString()
                                                  Dim Pitch As Integer = Integer.Parse(Me.Invoke(Function()
                                                                                                     Return TextBox1.Text
                                                                                                 End Function))
                                                  Dim Speed As Double = Double.Parse(Me.Invoke(Function()
                                                                                                   Return TextBox2.Text
                                                                                               End Function))
                                                  Dim compression As CompressionLevel = Me.Invoke(Function()
                                                                                                      Return ComboBox3.SelectedIndex
                                                                                                  End Function)
                                                  Dim display As Form2
                                                  Me.Invoke(Sub()
                                                                display = New Form2()
                                                                display.RichTextBox1.Text = text
                                                                display.Label1.Text = "Voice: " & Voice
                                                                display.Label2.Text = "Pitch: " & Pitch.ToString()
                                                                display.Label3.Text = "Speed: " & Speed.ToString()
                                                                display.Label4.Text = "Characters: " & text.Length.ToString()
                                                                display.Show()
                                                            End Sub)
                                                  Dim compressedText As String = Compress(text, compression)
                                                  If GetLength(compressedText, Voice, Pitch, Speed) Then
                                                      Me.Invoke(Sub()
                                                                    display.ProgressBar1.Value = 1000
                                                                End Sub)
                                                      Dim fileStream As IO.FileStream
                                                      While True
                                                          Try
                                                              fileStream = New IO.FileStream(destination, IO.FileMode.Create, IO.FileAccess.Write)
                                                              Exit While
                                                          Catch ex As IO.IOException
                                                              If MessageBox.Show("File " & destination & " is being used by another process. Terminate said process or try again later.", "Error", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error) = DialogResult.Cancel Then Return
                                                          End Try
                                                      End While
                                                      Try
                                                          Dim resultBytes As Byte() = If(Await GetAudio(compressedText, Voice, Pitch, Speed), {})
                                                          fileStream.Write(resultBytes, 0, resultBytes.Length)
                                                          fileStream.Close()
                                                      Catch ex As Exception
                                                      End Try
                                                      Me.Invoke(Sub()
                                                                    display.ProgressBar2.Value = 1000
                                                                End Sub)
                                                  Else
                                                      While text.Contains(newLineCharacter & newLineCharacter)
                                                          text = text.Replace(newLineCharacter & newLineCharacter, newLineCharacter)
                                                      End While
                                                      Dim textSubstrings As String() = text.Split(newLineCharacter)
                                                      Dim textSubstring As String = ""
                                                      Dim compressedTextSubstring As String = ""
                                                      Dim nextTextSubstring As String
                                                      Dim nextCompressedTextSubstring As String
                                                      For i As UInt64 = 0 To textSubstrings.Length - 1
                                                          nextTextSubstring = textSubstring & newLineCharacter & textSubstrings(i)
                                                          nextCompressedTextSubstring = Compress(nextTextSubstring, compression)
                                                          If (Not GetLength(nextCompressedTextSubstring, Voice, Pitch, Speed)) Or i + 1 = textSubstrings.Length Then
                                                              If textSubstring.Length = 0 Then
                                                                  MessageBox.Show("Text block too long. Consider adding returns to break up text.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                                                  Return
                                                              Else
                                                                  If i + 1 = textSubstring.Length Then compressedTextSubstring = nextCompressedTextSubstring
                                                                  If compressedTextSubstring.Length <> 0 Then
                                                                      strings.Add(compressedTextSubstring)
                                                                  End If
                                                                  nextTextSubstring = textSubstrings(i)
                                                                  nextCompressedTextSubstring = Compress(nextTextSubstring, compression)
                                                              End If
                                                          End If
                                                          textSubstring = nextTextSubstring
                                                          compressedTextSubstring = nextCompressedTextSubstring
                                                          Dim currentI As UInt64 = i
                                                          Me.Invoke(Sub()
                                                                        display.ProgressBar1.Value = Math.Round(1000 * currentI / textSubstrings.Length)
                                                                    End Sub)
                                                      Next
                                                      Me.Invoke(Sub()
                                                                    display.ProgressBar1.Value = 1000
                                                                End Sub)
                                                      Dim files(strings.Count - 1)
                                                      Dim fileStreams(strings.Count - 1) As IO.FileStream
                                                      For i As UInt64 = 0 To files.Length - 1
                                                          files(i) = destination.Substring(0, destination.Length - destination.Split("."c)(destination.Split("."c).Length - 1).Length - 1) & " - Part " & (i + 1) & "." & destination.Split("."c)(destination.Split("."c).Length - 1)
                                                          While True
                                                              Try
                                                                  fileStreams(i) = New IO.FileStream(files(i), IO.FileMode.Create, IO.FileAccess.Write)
                                                                  Exit While
                                                              Catch ex As IO.IOException
                                                                  If MessageBox.Show("File " & files(i) & " is being used by another process. Terminate said process or try again later.", "Error", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error) = DialogResult.Cancel Then Return
                                                              End Try
                                                          End While
                                                      Next
                                                      For i As Integer = 0 To strings.Count - 1
                                                          While Not Me.Invoke(Function()
                                                                                  If (Date.Now - lastRequestTime).TotalMinutes > 0 Then
                                                                                      lastRequestTime = Date.Now
                                                                                      requestCount = 0
                                                                                  End If
                                                                                  If requestCount >= RequestRate Then Return False
                                                                                  requestCount += 1
                                                                                  Return True
                                                                              End Function)
                                                              Threading.Thread.Sleep(1000)
                                                          End While
                                                          Try
                                                              Dim resultBytes As Byte() = Await GetAudio(strings(i), Voice, Pitch, Speed)
                                                              If IsNothing(resultBytes) Then
                                                                  If Not IO.File.Exists(files(i)) Then IO.File.WriteAllText(files(i), strings(i))
                                                              Else
                                                                  fileStreams(i).Write(resultBytes, 0, resultBytes.Length)
                                                              End If
                                                              fileStreams(i).Close()
                                                          Catch ex As Exception
                                                          End Try
                                                          Dim currentI As Integer = i
                                                          Me.Invoke(Sub()
                                                                        display.ProgressBar2.Value = Math.Round(1000 * (currentI + 1) / strings.Count)
                                                                    End Sub)
                                                      Next
                                                      Me.Invoke(Sub()
                                                                    display.ProgressBar2.Value = 1000
                                                                End Sub)
                                                  End If
                                                  Me.Invoke(Sub()
                                                                display.Dispose()
                                                                processes -= 1
                                                            End Sub)
                                              End Sub)
        newThread.Start()
    End Sub
    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If processes > 0 Then e.Cancel = (MessageBox.Show("There " & If(processes = 1, "is", "are") & " still " & processes & If(processes = 1, " process", " processes") & " running. Cancel anyway?", "Termination confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.No)
    End Sub
End Class
Public Class ResponseType0
    Public audioContent As String
End Class
Public Enum CompressionLevel
    Level0 = 0
    Level1 = 1
    Level2 = 2
End Enum
Public Class RawVoiceInformation
    Public languageCodes As String()
    Public name As String
    Public ssmlGender As String
    Public naturalSampleRateHertz As UInt64
End Class
Public Enum SsmlGender
    female
    male
End Enum
Public Class ResponseType1
    Public voices As RawVoiceInformation()
End Class