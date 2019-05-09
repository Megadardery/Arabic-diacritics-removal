Module mdlMain
    Sub Main()
        'If False Then
        '    Debugging()
        '    Exit Sub
        'End If

        Try
            Console.WriteLine("Diacritics removal started.")
            Dim str As String = My.Computer.Clipboard.GetText()
            If str = "" Then GoTo Finish
            Dim newstr As New Text.StringBuilder
            For Each Chr As Char In str
                newstr.Append(Valify(Chr))
            Next
            If newstr.ToString = "" Then My.Computer.Clipboard.Clear() Else My.Computer.Clipboard.SetText(newstr.ToString)
Finish:     Console.WriteLine("Diacritics removal ended.")


        Catch ex As Exception
            Console.WriteLine("Diacritics removal halted due to an error.")
            Console.WriteLine(ex.Message)
            Console.WriteLine()
            Console.WriteLine("Press any key to exit.")
            Console.ReadKey()
        End Try
    End Sub

    Function Valify(chr As Char) As String
        Select Case AscW(chr)
            Case &H623 To &H63A, &H641 To &H64A         'ا ب ة ت ث ج ح خ د ذ ر ز س ش ص ض ط ظ ع غ  'ف ق ك ل م ن ه و ى ي
                Return chr
            Case &H7E                                   '~
                Return ""
            Case &H8E4, &H610 To &H61A, &H640, &H64B To &H65F   'Ordinary diacritics
                Return ""
            Case &H620, &H63D To &H63F, &H6CC To &H6D2
                Return "ي"
            Case &H622, &H670 To &H671
                Return "ا"
            Case &H675
                Return "أ"
            Case &H676 To &H677
                Return "ؤ"
            Case &H678, &H6D3
                Return "ئ"
            Case &H63B To &H63C, &H762 To &H764, &H6A9 To &H6B4
                Return "ك"
            Case &H686
                Return "ج"
            Case &H688 To &H690
                Return "د"
            Case &H691 To &H699
                Return "ر"
            Case &H6B5 To &H6B8
                Return "ل"
            Case &H6BE, &H6C0 To &H6C2, &H6D5
                Return "ه"
            Case &H6C3
                Return "ة"
            Case &H6C4 To &H6CB
                Return "و"
            Case Else
                Return chr
        End Select
    End Function

    'Sub Debugging()
    '    Console.WriteLine("The application is currently in debugging mode. Press any key to start debugging the text in clipboard.")
    '    Console.Read()
    '    Dim str As String = My.Computer.Clipboard.GetText()
    '    Console.WriteLine("-----------------------------------------")

    '    For Each character As Char In str
    '        Dim newstr As New Text.StringBuilder
    '        newstr.Append(character)
    '        newstr.Append(vbTab)
    '        newstr.Append(vbTab)
    '        newstr.Append(vbTab)
    '        newstr.Append("|")
    '        newstr.Append(vbTab)
    '        newstr.Append(vbTab)
    '        newstr.Append("&H")
    '        newstr.Append(Hex(AscW(character)))
    '        Console.WriteLine(newstr.ToString)
    '    Next
    '    Console.WriteLine("-----------------------------------------")
    '    Console.WriteLine("Press any key to exit.")
    '    Console.ReadKey()
    'End Sub
End Module

