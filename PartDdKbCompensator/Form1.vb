Public Class Form1
    'v1.0.1 3/23/24Sa 730pm: Lazy way, hard coded
    'v1.0.2 3/23/24Sa 836pm: 1st part of the customizable substitutions way, a Subst() 40-max-element array, that will store
    '    all the key substitutions, both subbee and subber. Tested, works. Kp it commented 4 now.
    'v1.0.3 3/23/24Sa 844pm: Making it more real instead of just testing gud enuf; erase subbee chars, only keep the added subber
    'v1.0.4 3/23/24Sa 946pm: After a test run, typed a whole essay in it, debugged couple key loose ends. SelStart, etc. dlz. OK and
    '    can type at the non-end which works but puts SelStart ONE char too far ahead then. Probly EZ fix. OK it was. Rmv "+ 1"!
    '    was only working b4 cuz SelSt of Len(Text1) + 1 just defaults to being AT the end, no error trigg'd, just givs u at-the-
    '    end and no further. And named the application title bar with 'partiall ded kb sol'n'--maybe refine that name but if can't
    '    think of els, or no time to, then there u go. First actally compiling an exe. I put rev comment in the comment, but to
    '    pretend this is a real application, I'll just put a general comment there that explains what the application does, and
    '    keep the rev notes here in this code comments area. OK? OK.
    Dim BuffSomTyping As String

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Dim Subst(40) As String '4 L8r
        Subst(1) = "][,h" 'To test at first til a user interface for this
        Subst(2) = "}{,H" 'To test at first til a user interface for this
        Subst(3) = ",.,u" 'To test at first til a user interface for this
        Dim OrigSelSt As Object
        Dim DontClr As Boolean

        OrigSelSt = TextBox1.SelectionStart

        Debug.Print(Asc(e.KeyChar))

        If Asc(e.KeyChar) = 9 Then 'This doesn't recognize tab in VS!
            TextBox1.Text = TextBox1.Text + "    "
        Else
            BuffSomTyping = BuffSomTyping & e.KeyChar
        End If
        If Strings.Right(BuffSomTyping, 2) = "][" Then 'h
            TextBox1.Text = Strings.Left(TextBox1.Text, Len(TextBox1.Text) - 1) & "h"
        ElseIf Strings.Right(BuffSomTyping, 2) = "}{" Then 'H
            TextBox1.Text = Strings.Left(TextBox1.Text, Len(TextBox1.Text) - 1) & "H"
        ElseIf Strings.Right(BuffSomTyping, 2) = ",." Then 'u
            TextBox1.Text = Strings.Left(TextBox1.Text, Len(TextBox1.Text) - 1) & "u"
        ElseIf Strings.Right(BuffSomTyping, 2) = "<>" Then 'U
            TextBox1.Text = Strings.Left(TextBox1.Text, Len(TextBox1.Text) - 1) & "U"
        ElseIf Strings.Right(BuffSomTyping, 2) = ",," Then 'j
            TextBox1.Text = Strings.Left(TextBox1.Text, Len(TextBox1.Text) - 1) & "j"
        ElseIf Strings.Right(BuffSomTyping, 2) = "<<" Then 'J
            TextBox1.Text = Strings.Left(TextBox1.Text, Len(TextBox1.Text) - 1) & "J"
        ElseIf Strings.Right(BuffSomTyping, 2) = ".," Then ';
            TextBox1.Text = Strings.Left(TextBox1.Text, Len(TextBox1.Text) - 1) & ";"
        ElseIf Strings.Right(BuffSomTyping, 2) = "><" Then ':
            TextBox1.Text = Strings.Left(TextBox1.Text, Len(TextBox1.Text) - 1) & ":"
        ElseIf Strings.Right(BuffSomTyping, 2) = "!x" Then '6
            TextBox1.Text = Strings.Left(TextBox1.Text, Len(TextBox1.Text) - 1) & "6"
        ElseIf Strings.Right(BuffSomTyping, 2) = "!X" Then '^
            TextBox1.Text = Strings.Left(TextBox1.Text, Len(TextBox1.Text) - 1) & "^"
        ElseIf Strings.Right(BuffSomTyping, 2) = "!v" Then '7
            TextBox1.Text = Strings.Left(TextBox1.Text, Len(TextBox1.Text) - 1) & "7"
        ElseIf Strings.Right(BuffSomTyping, 2) = "!V" Then '&
            TextBox1.Text = Strings.Left(TextBox1.Text, Len(TextBox1.Text) - 1) & "&"
        ElseIf Strings.Right(BuffSomTyping, 3) = "/'" Then '\  -pref'd subbee is /// but nd timer for doing // in that case, timeout on
            TextBox1.Text = Strings.Left(TextBox1.Text, Len(TextBox1.Text) - 1) & "\"                      '                     // at end of the buffer means barf a w and forget about
        ElseIf Strings.Right(BuffSomTyping, 2) = "?'" Then '|                    the \ idea. type the 3rd / before timeout and get ur \
            TextBox1.Text = Strings.Left(TextBox1.Text, Len(TextBox1.Text) - 1) & "|"
        ElseIf Strings.Right(BuffSomTyping, 2) = "//" Then 'w
            'TextBox1.Text = Strings.Left(TextBox1.Text, Len(TextBox1.Text) - 1) & "w"
            TextBox1.Text = Strings.Left(TextBox1.Text, OrigSelSt - 1) & "w" & Strings.Right(TextBox1.Text, Len(TextBox1.Text) - OrigSelSt)
        ElseIf Strings.Right(BuffSomTyping, 2) = "??" Then 'W
            TextBox1.Text = Strings.Left(TextBox1.Text, Len(TextBox1.Text) - 1) & "W"
        Else
            DontClr = True
        End If

        '    For i = 1 To 40
        '        If Not Subst(i) = "" Then
        '
        '            If Strings.Right(BuffSomTyping, 2) = Strings.Left(Subst(i), 2) Then
        '                TextBox1.Text = TextBox1.Text & Strings.Right(Subst(i), 1)
        '                Exit For
        '            End If
        '
        '        End If
        '    Next
        '    If i >= 40 Then
        '       DontClr = True
        '    End If
        'It WORKS! Abov "v1" way gud enuf 4 now tho

        If Not DontClr Then 'That is, if we typed a subbee, 2nd OR 1st char? No, just upon 2nd char of, m?
            e.KeyChar = "_" 'Starting get serious, no more just testing the concept; remove subbee chars, cuz putting subber instead.
            BuffSomTyping = ""
            '        TextBox1..SelectionStart = Len(TextBox1.Text) '2nd part of the 2-lines-above deal, SelectionStart got 1 off, maybe doesn't happen anymore
            'but wutev, ensure we at end of what we're typing. AHA! Tx.SelSt = Tx.SelSt + (or -) 1 can mk this useable ANYwhere, COO.
            TextBox1.SelectionStart = OrigSelSt '+ 1 'See abov's comment
        End If
        '^^^if u nev typ a substitute it'll ovflow but low pri concern

        'TODO-- user customizable 'substitutes', click btn, goes to a scrn 4 that, sprdshtish rm for 40.
        '][ for h changed to )( instead? Subst(1) from "][,h" to ")(,h", OK? Mk up the better datastruc or
        'Dictionary etc. way to do l8r, even just a 2D array w/ "][" in subst(1,1) and "h" in subst(1,2) and
        'on like that. nameable cols instd? col1 is subber and col2 is subee, I'd say.

    End Sub
End Class
