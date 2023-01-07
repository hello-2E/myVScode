Imports System.IO
Imports Spire.Doc
Module Module1
    Public lines As String()
    Public arr
    Sub Main()
        Dim m As Integer = -1
        Dim myr
        lines = File.ReadAllLines("C:\Users\2E\Desktop\A.txt")
        Dim d As Dictionary(Of String, String) = New Dictionary(Of String, String)
        ReDim arr(lines.Length - 1, 9)
        For I = 1 To lines.Length - 1
            Dim SR = Split(lines(I), vbTab)
            If Not d.ContainsKey(SR(3)) Then
                m = m + 1
                d(SR(3)) = m
                arr(m, 0) = SR(3) ''户号
                If SR(9) = "本人" Then
                    arr(m, 1) = SR(5) ''户主姓名
                    arr(m, 2) = SR(6) ''户主性别
                    arr(m, 3) = SR(7) ''户主民族
                    arr(m, 4) = SR(9) ''与户主关系
                    arr(m, 5) = SR(10) ''家庭地址
                    arr(m, 9) = SR(8)
                End If
                arr(m, 6) = 1 ''户数
                arr(m, 7) = SR(2) ''经济组织名称
                arr(m, 8) = I
            Else
                myr = d(SR(3))
                If SR(9) = "户主" Then
                    arr(myr, 1) = SR(5) ''户主姓名
                    arr(myr, 2) = SR(6) ''户主性别
                    arr(myr, 3) = SR(7) ''户主民族
                    arr(myr, 4) = SR(9) ''与户主关系
                    arr(myr, 5) = SR(10) ''家庭地址
                    arr(myr, 9) = SR(8)
                End If
                arr(myr, 6) = arr(myr, 6) + 1
                arr(myr, 8) = arr(myr, 8) & "," & I
            End If
        Next
        Dim c As Integer = m / 7
        Dim te1 = New System.Threading.Thread(AddressOf test1)
        Dim te2 = New System.Threading.Thread(AddressOf test1)
        Dim te3 = New System.Threading.Thread(AddressOf test1)
        Dim te4 = New System.Threading.Thread(AddressOf test1)
        Dim te5 = New System.Threading.Thread(AddressOf test1)
        Dim te6 = New System.Threading.Thread(AddressOf test1)
        Dim te7 = New System.Threading.Thread(AddressOf test1)
        Dim str1 As String = "0" & "," & c & ",1"
        Dim str2 As String = c + 1 & "," & c * 2 & ",2"
        Dim str3 As String = c * 2 + 1 & "," & c * 3 & ",3"
        Dim str4 As String = c * 3 + 1 & "," & c * 4 & ",4"
        Dim str5 As String = c * 4 + 1 & "," & c * 5 & ",5"
        Dim str6 As String = c * 5 + 1 & "," & c * 6 & ",6"
        Dim str7 As String = c * 6 + 1 & "," & m & ",7"
        te1.Start(str1)
        te2.Start(str2)
        te3.Start(str3)
        te4.Start(str4)
        te5.Start(str5）
        te6.Start(str6)
        te7.Start(str7)
    End Sub
    Private Function Boo_DirExist(ByVal Str_Path As String) As Boolean
        Boo_DirExist = System.IO.Directory.Exists(Str_Path)
    End Function
    'Private Sub CreateWb()
    '    For i = 1 To 1000
    '        Dim wDoc As XWPFDocument = Nothing
    '        Dim filePath As String = "C:\Users\2E\Desktop\PDF\Mo.docx"
    '        Using fs As New FileStream(filePath, FileMode.Open)
    '            wDoc = New XWPFDocument(fs)
    '        End Using
    '        Dim str As String = "第A号"
    '        For Each prg As XWPFParagraph In wDoc.Paragraphs
    '            If InStr(prg.Text, "第A") > 0 Then
    '                prg.ReplaceText(str, i) '''替换
    '                Exit For
    '                Stop
    '            End If
    '        Next
    '        Dim fileS As New FileStream("C:\Users\2E\Desktop\PDF\A\" & i & ".docx", FileMode.Create)
    '        wDoc.Write(fileS)
    '        wDoc.Close()
    '    Next

    'End Sub
    Private Sub getPDF()
        Dim document As New Spire.Doc.Document

        document.LoadFromFile("C:\Users\2E\Desktop\PDF\Mo.docx", FileFormat.Docx)
        For i = 1 To 100
            document.Replace("第A号", "第00001号", False, True)
            Dim section = document.Sections(0)
            Dim table = section.Tables(0)
            Dim cell = table.Rows(1).Cells(1)
            Dim p1 = cell.Paragraphs(0)
            p1.Text = i
            'document.Sections(0).Tables(0).Rows(1).Cells(1)


            'Console.WriteLine(dp)
            document.SaveToFile("C:\Users\2E\Desktop\PDF\A\" & i & ".pdf", fileFormat:=FileFormat.PDF)
        Next
        document.Close()

    End Sub
    Private Sub test1(str)
        Dim spath As String = ""
        Dim document As New Spire.Doc.Document
        Dim n As Integer
        Dim s As Integer = Split(str, ",")(0)
        Dim e As Integer = Split(str, ",")(1)
        Dim x As Integer = Split(str, ",")(2)
        For ii = s To e '00
            document.LoadFromFile("C:\Users\2E\Desktop\PDF\" & x & ".docx", FileFormat.Docx)
            document.Replace("第A号", arr(ii, 0) & "号", False, True)
            Dim section = document.Sections(0)
            Dim table = section.Tables(0)
            Dim c0 = table.Rows(0).Cells(1)
            Dim p0 = c0.Paragraphs(0)
            p0.Text = arr(ii, 7)

            Dim c1 = table.Rows(1).Cells(1)
            Dim p1 = c1.Paragraphs(0)
            p1.Text = arr(ii, 1)

            Dim c2 = table.Rows(1).Cells(3)
            Dim p2 = c2.Paragraphs(0)
            p2.Text = arr(ii, 2)


            Dim c3 = table.Rows(1).Cells(5)
            Dim p3 = c3.Paragraphs(0)
            p3.Text = arr(ii, 3)


            Dim c4 = table.Rows(1).Cells(7)
            Dim p4 = c4.Paragraphs(0)
            p4.Text = arr(ii, 6)

            Dim c5 = table.Rows(2).Cells(1)
            Dim p5 = c5.Paragraphs(0)
            p5.Text = arr(ii, 9)

            Dim c6 = table.Rows(2).Cells(3)
            Dim p6 = c6.Paragraphs(0)
            p6.Text = arr(ii, 6)

            Dim c7 = table.Rows(3).Cells(1)
            Dim p7 = c7.Paragraphs(0)
            p7.Text = arr(ii, 5)
            n = 6
            Dim ar = Split(arr(ii, 8), ",")
            For i2 = 0 To UBound(ar)
                n = n + 1
                Dim cel1 = table.Rows(n).Cells(0)
                Dim cel2 = table.Rows(n).Cells(1)
                Dim cel3 = table.Rows(n).Cells(2)
                Dim cel4 = table.Rows(n).Cells(3)
                Dim A1 = cel1.Paragraphs(0)
                Dim A2 = cel2.Paragraphs(0)
                Dim A3 = cel3.Paragraphs(0)
                Dim A4 = cel4.Paragraphs(0)
                Dim sr2 = Split(lines(Val(ar(i2))), vbTab)
                If spath = "" Then
                    spath = "C:\Users\2E\Desktop\PDF\result\" & sr2(0) & "\" & sr2(1) & "\"
                    If Not Directory.Exists(spath) Then
                        Directory.CreateDirectory(spath)
                    End If
                End If
                A1.Text = sr2(5)
                A2.Text = sr2(9)
                A3.Text = sr2(8)
                A4.Text = 1
            Next
            document.SaveToFile(spath & arr(ii, 0) & arr(ii, 1) & ".pdf", fileFormat:=FileFormat.PDF)
            document.Close()
            spath = ""
        Next
    End Sub
End Module
