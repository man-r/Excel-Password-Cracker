Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Public Class Form1

    Private Sub Browse_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Browse_Button.Click
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            FilePath_TextBox.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        If 1 > 0 Then
            manarCrackerCode()
            Exit Sub
        End If

        
        
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim oXl As Excel.Application
        Dim oWb As Excel.Workbook
        Dim oWs As Excel.Worksheet
        'Dim chart As Excel.Chart

        oXl = CreateObject("Excel.Application")
        oWb = oXl.Workbooks.Open(FilePath_TextBox.Text)

        oXl.Visible = True

        For Each oWs In oWb.Worksheets
            Try
                oWs.Visible = Excel.XlSheetVisibility.xlSheetVisible
            Catch ex As Exception
                MsgBox("eRROR tRYING TO uNHIDE sHEET "" " & oWs.Name & " "" " & vbNewLine & vbNewLine & "eRROR: " & ex.Message & vbNewLine & vbNewLine & "tHE fILE mIGHT bE pROTECTED iF sO uNPROTECTED fIRST")
            End Try

        Next

        oWs = Nothing
        oWb = Nothing
        oXl = Nothing
    End Sub


    Function toBinary(ByVal IntNum As Integer) As String


        Dim TempValue As Integer
        
        Dim str As String = ""

        Do
            'Use the Mod operator to get the current binary digit from the
            'Integer number
            TempValue = IntNum Mod 2
            TempValue = TempValue + 65

            str = Chr(TempValue) + str

            'Divide the current number by 2 and get the integer result
            IntNum = IntNum \ 2
        Loop Until IntNum = 0

        While str.Length < 11
            str = "A" & str
        End While

        Return str
    End Function


    Sub manarCrackerCode()
        Dim oXl As Excel.Application
        Dim oWb As Excel.Workbook
        Dim oWs As Excel.Worksheet
        'Dim chart As Excel.Chart

        oXl = CreateObject("Excel.Application")
        oWb = oXl.Workbooks.Open(FilePath_TextBox.Text)
        oWs = oWb.ActiveSheet
        oXl.Visible = True

        'Breaks worksheet and workbook structure passwords.
        'Bob McCormick probably originator of base code algorithm
        'Modified for coverage of workbook structure / windows
        'passwords and for multiple passwords
        'Norman Harker and JE McGimpsey 27-Dec-2002
        'Reveals passwords NOT "the" passwords

        Dim str As String
        Dim Mess As String
        Dim Header As String
        Dim AllClear As String
        Dim PWord1 As String = ""
        Dim ShTag As Boolean
        Dim WinTag As Boolean
        Dim w1 As Worksheet
        Dim w2 As Worksheet
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim m As Integer
        Dim n As Integer
        Dim i1 As Integer
        Dim i2 As Integer
        Dim i3 As Integer
        Dim i4 As Integer
        Dim i5 As Integer
        Dim i6 As Integer


        'oXl.ScreenUpdating = False
        Header = "manar excel password cracker"
        AllClear = vbNewLine & vbNewLine & "The workbook should now be free of all password protection so " & _
        "make sure you:" & vbNewLine & vbNewLine & "SAVE IT NOW!" & vbNewLine & vbNewLine & "and also" & vbNewLine & vbNewLine & "BACKUP!, BACKUP!!, BACKUP!!!" & vbNewLine & vbNewLine & _
        "Also, remember that the password " & _
        "was put there for a reason. Don't " & _
        "stuff up crucial formulas or data." & _
        vbNewLine & vbNewLine & "Access and use of some data may" & _
        "be an offence. If in doubt, don't."

        With oWb
            WinTag = .ProtectStructure Or .ProtectWindows
        End With

        ShTag = False

        For Each w1 In oWb.Worksheets
            ShTag = ShTag Or w1.ProtectContents
        Next w1

        If Not ShTag And Not WinTag Then
            MsgBox("There were no passwords on sheets, or workbook structure or windows.")
            Exit Sub
        End If

        Mess = "After pressing OK button this will take some time." & _
        vbNewLine & vbNewLine & "Amount of time depends on how " & _
        "many different passwords, the passwords, and" & _
        "your computer's specification." & vbNewLine & vbNewLine & _
        "Just be patient! Make me a coffee!"

        MsgBox(Mess, vbInformation, Header)



        If Not WinTag Then ' no protection on the excel file
            Mess = "There was no protection to workbook structure " & _
            "or windows." & vbNewLine & vbNewLine & _
            "Proceeding to unprotect sheets."
            MsgBox(Mess, vbInformation, Header)

        Else
            On Error Resume Next
            Do 'dummy do loop

                For i = 0 To 2047
                    str = toBinary(i)
                    For n = 32 To 126
                        With oWb
                            .Unprotect(str & Chr(n))

                            Label1.Text = str & Chr(n)

                            If .ProtectStructure = False And _
                            .ProtectWindows = False Then
                                PWord1 = str & Chr(n)
                                Mess = "You had a Worksheet Structure or " & _
                                "Windows Password set." & vbNewLine & vbNewLine & _
                                "The password found was: " & vbNewLine & vbNewLine & _
                                PWord1 & vbNewLine & vbNewLine & "Note it down for " & _
                                "potential future use in other " & _
                                "workbooks by same person who set this " & _
                                "password." & vbNewLine & vbNewLine & _
                                "Now to check and clear other passwords."
                                MsgBox(Mess, vbInformation, Header)
                                Exit Do 'Bypass all for...nexts
                            End If
                        End With
                    Next : Next
            Loop Until True
            On Error GoTo 0
        End If

        If WinTag And Not ShTag Then
            Mess = "Only structure / windows protected with " & _
            "the password that was just found." & _
            AllClear
            MsgBox(Mess, vbInformation, Header)
            Exit Sub
        End If

        On Error Resume Next

        For Each w1 In oWb.Worksheets
            'Attempt clearance with PWord1
            w1.Unprotect(PWord1)
        Next w1
        On Error GoTo 0

        ShTag = False

        For Each w1 In oWb.Worksheets
            'Checks for all clear ShTag triggered to 1 if not.
            ShTag = ShTag Or w1.ProtectContents
        Next w1

        If Not ShTag Then
            Mess = AllClear
            MsgBox(Mess, vbInformation, Header)
            Exit Sub
        End If

        For Each w1 In oWb.Worksheets
            With w1

                If .ProtectContents And MsgBox("trying to crack " & w1.Name & " sheet password" & vbNewLine & vbNewLine & "DO YOU WHANT TO CONTINUE?", vbYesNo) = vbYes Then

                    On Error Resume Next

                    Do 'Dummy do loop
                        For i = 65 To 66 : For j = 65 To 66 : For k = 65 To 66
                                    For l = 65 To 66 : For m = 65 To 66 : For i1 = 65 To 66
                                                For i2 = 65 To 66 : For i3 = 65 To 66 : For i4 = 65 To 66
                                                            For i5 = 65 To 66 : For i6 = 65 To 66 : For n = 32 To 126
                                                                        .Unprotect(Chr(i) & Chr(j) & Chr(k) & _
                                                                        Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
                                                                        Chr(i4) & Chr(i5) & Chr(i6) & Chr(n))

                                                                        Label1.Text = Chr(i) & Chr(j) & Chr(k) & _
                                                                    Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
                                                                    Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)


                                                                        If Not .ProtectContents Then

                                                                            PWord1 = Chr(i) & Chr(j) & Chr(k) & Chr(l) & _
                                                                            Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
                                                                            Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
                                                                            Mess = "You had a Worksheet password set." & _
                                                                            vbNewLine & vbNewLine & "The password found was: " & _
                                                                            vbNewLine & vbNewLine & PWord1 & vbNewLine & vbNewLine & _
                                                                            "Note it down for potential future use " & _
                                                                            "in other workbooks by same person who " & _
                                                                            "set this password." & vbNewLine & vbNewLine & _
                                                                            "Now to check and clear other passwords."
                                                                            MsgBox(Mess, vbInformation, Header)

                                                                            'leverage finding Pword by trying on other sheets
                                                                            For Each w2 In oWb.Worksheets
                                                                                w2.Unprotect(PWord1)
                                                                            Next (w2)

                                                                            Exit Do 'Bypass all for...nexts
                                                                        End If

                                                                    Next : Next : Next : Next : Next : Next
                                            Next : Next : Next : Next : Next : Next
                    Loop Until True
                    On Error GoTo 0
                End If
            End With
        Next w1

        oWs = Nothing
        oWb = Nothing
        oXl = Nothing

    End Sub

    Sub InternetCracker()
        Dim oXl As Excel.Application
        Dim oWb As Excel.Workbook
        Dim oWs As Excel.Worksheet
        'Dim chart As Excel.Chart

        oXl = CreateObject("Excel.Application")
        oWb = oXl.Workbooks.Open(FilePath_TextBox.Text)
        oWs = oWb.ActiveSheet
        oXl.Visible = True

        'Breaks worksheet and workbook structure passwords.
        'Bob McCormick probably originator of base code algorithm
        'Modified for coverage of workbook structure / windows
        'passwords and for multiple passwords
        'Norman Harker and JE McGimpsey 27-Dec-2002
        'Reveals passwords NOT "the" passwords

        Dim Mess As String
        Dim Header As String
        Dim AllClear As String
        Dim PWord1 As String = ""
        Dim ShTag As Boolean
        Dim WinTag As Boolean
        Dim w1 As Worksheet
        Dim w2 As Worksheet
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim m As Integer
        Dim n As Integer
        Dim i1 As Integer
        Dim i2 As Integer
        Dim i3 As Integer
        Dim i4 As Integer
        Dim i5 As Integer
        Dim i6 As Integer


        'oXl.ScreenUpdating = False
        Header = "manar excel password cracker"
        AllClear = vbNewLine & vbNewLine & "The workbook should now be free of all password protection so " & _
        "make sure you:" & vbNewLine & vbNewLine & "SAVE IT NOW!" & vbNewLine & vbNewLine & "and also" & vbNewLine & vbNewLine & "BACKUP!, BACKUP!!, BACKUP!!!" & vbNewLine & vbNewLine & _
        "Also, remember that the password " & _
        "was put there for a reason. Don't " & _
        "stuff up crucial formulas or data." & _
        vbNewLine & vbNewLine & "Access and use of some data may" & _
        "be an offence. If in doubt, don't."

        With oWb
            WinTag = .ProtectStructure Or .ProtectWindows
        End With

        ShTag = False

        For Each w1 In oWb.Worksheets
            ShTag = ShTag Or w1.ProtectContents
        Next w1

        If Not ShTag And Not WinTag Then
            MsgBox("There were no passwords on sheets, or workbook structure or windows.")
            Exit Sub
        End If

        Mess = "After pressing OK button this will take some time." & _
        vbNewLine & vbNewLine & "Amount of time depends on how " & _
        "many different passwords, the passwords, and" & _
        "your computer's specification." & vbNewLine & vbNewLine & _
        "Just be patient! Make me a coffee!"

        MsgBox(Mess, vbInformation, Header)



        If Not WinTag Then ' no protection on the excel file
            Mess = "There was no protection to workbook structure " & _
            "or windows." & vbNewLine & vbNewLine & _
            "Proceeding to unprotect sheets."
            MsgBox(Mess, vbInformation, Header)

        Else
            On Error Resume Next
            Do 'dummy do loop

                For i = 65 To 66 : For j = 65 To 66 : For k = 65 To 66
                            For l = 65 To 66 : For m = 65 To 66 : For i1 = 65 To 66
                                        For i2 = 65 To 66 : For i3 = 65 To 66 : For i4 = 65 To 66
                                                    For i5 = 65 To 66 : For i6 = 65 To 66 : For n = 32 To 126
                                                                With oWb
                                                                    .Unprotect(Chr(i) & Chr(j) & Chr(k) & _
                                                                    Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
                                                                    Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n))

                                                                    Label1.Text = Chr(i) & Chr(j) & Chr(k) & _
                                                                    Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
                                                                    Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)

                                                                    If .ProtectStructure = False And _
                                                                    .ProtectWindows = False Then
                                                                        PWord1 = Chr(i) & Chr(j) & Chr(k) & Chr(l) & _
                                                                        Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
                                                                        Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
                                                                        Mess = "You had a Worksheet Structure or " & _
                                                                        "Windows Password set." & vbNewLine & vbNewLine & _
                                                                        "The password found was: " & vbNewLine & vbNewLine & _
                                                                        PWord1 & vbNewLine & vbNewLine & "Note it down for " & _
                                                                        "potential future use in other " & _
                                                                        "workbooks by same person who set this " & _
                                                                        "password." & vbNewLine & vbNewLine & _
                                                                        "Now to check and clear other passwords."
                                                                        MsgBox(Mess, vbInformation, Header)
                                                                        Exit Do 'Bypass all for...nexts
                                                                    End If
                                                                End With
                                                            Next : Next : Next : Next : Next : Next
                                    Next : Next : Next : Next : Next : Next

            Loop Until True
            On Error GoTo 0
        End If

        If WinTag And Not ShTag Then
            Mess = "Only structure / windows protected with " & _
            "the password that was just found." & _
            AllClear
            MsgBox(Mess, vbInformation, Header)
            Exit Sub
        End If

        On Error Resume Next

        For Each w1 In oWb.Worksheets
            'Attempt clearance with PWord1
            w1.Unprotect(PWord1)
        Next w1
        On Error GoTo 0

        ShTag = False

        For Each w1 In oWb.Worksheets
            'Checks for all clear ShTag triggered to 1 if not.
            ShTag = ShTag Or w1.ProtectContents
        Next w1

        If Not ShTag Then
            Mess = AllClear
            MsgBox(Mess, vbInformation, Header)
            Exit Sub
        End If

        For Each w1 In oWb.Worksheets
            With w1

                If .ProtectContents And MsgBox("trying to crack " & w1.Name & " sheet password" & vbNewLine & vbNewLine & "DO YOU WHANT TO CONTINUE?", vbYesNo) = vbYes Then

                    On Error Resume Next

                    Do 'Dummy do loop
                        For i = 65 To 66 : For j = 65 To 66 : For k = 65 To 66
                                    For l = 65 To 66 : For m = 65 To 66 : For i1 = 65 To 66
                                                For i2 = 65 To 66 : For i3 = 65 To 66 : For i4 = 65 To 66
                                                            For i5 = 65 To 66 : For i6 = 65 To 66 : For n = 32 To 126
                                                                        .Unprotect(Chr(i) & Chr(j) & Chr(k) & _
                                                                        Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
                                                                        Chr(i4) & Chr(i5) & Chr(i6) & Chr(n))

                                                                        Label1.Text = Chr(i) & Chr(j) & Chr(k) & _
                                                                    Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
                                                                    Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)


                                                                        If Not .ProtectContents Then

                                                                            PWord1 = Chr(i) & Chr(j) & Chr(k) & Chr(l) & _
                                                                            Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
                                                                            Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
                                                                            Mess = "You had a Worksheet password set." & _
                                                                            vbNewLine & vbNewLine & "The password found was: " & _
                                                                            vbNewLine & vbNewLine & PWord1 & vbNewLine & vbNewLine & _
                                                                            "Note it down for potential future use " & _
                                                                            "in other workbooks by same person who " & _
                                                                            "set this password." & vbNewLine & vbNewLine & _
                                                                            "Now to check and clear other passwords."
                                                                            MsgBox(Mess, vbInformation, Header)

                                                                            'leverage finding Pword by trying on other sheets
                                                                            For Each w2 In oWb.Worksheets
                                                                                w2.Unprotect(PWord1)
                                                                            Next (w2)

                                                                            Exit Do 'Bypass all for...nexts
                                                                        End If

                                                                    Next : Next : Next : Next : Next : Next
                                            Next : Next : Next : Next : Next : Next
                    Loop Until True
                    On Error GoTo 0
                End If
            End With
        Next w1

        oWs = Nothing
        oWb = Nothing
        oXl = Nothing
    End Sub
End Class
