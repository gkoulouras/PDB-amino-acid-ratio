Imports OfficeOpenXml
Imports System.Drawing
Imports System.IO

Module Module1

    Sub Main()

        Dim pck As New ExcelPackage
        Dim ws As ExcelWorksheet
        Dim stpw As New Stopwatch
        'Dim path As String = "C:\Users\grigo\Dropbox\PROJECTS\PEZ_THESIS\output_dimers"
        Dim fileContents As New System.Text.StringBuilder()
        Dim fileContent As String = String.Empty
        Dim line As String = String.Empty
        Dim linecounter As Integer = 0
        Dim filecounter As Integer = 0
        Dim generalcounter, Acounter, Ccounter, Dcounter, Ecounter, Fcounter, Gcounter, Hcounter, Icounter,
        Kcounter, Lcounter, Mcounter, Ncounter, Pcounter, Qcounter, Rcounter, Scounter,
        Tcounter, Vcounter, Wcounter, Ycounter As Integer

        Try

            Console.WriteLine("The application started..." & vbCrLf)
            Console.WriteLine("Please declare the full path of the folder which contains the PDB files:" & vbCrLf)
            Dim path As String = Console.ReadLine
            stpw.Restart()
            Console.WriteLine("The following files have been retrieved:" & vbCrLf)
            For Each f As FileInfo In New DirectoryInfo(path).GetFiles("*.pdb") ' Specify a file pattern here

                Dim sr As New IO.StreamReader(f.FullName)

                Do While Not sr.EndOfStream
                    line = sr.ReadLine
                    linecounter = linecounter + 1
                    'Console.WriteLine(f.ToString & " - " & line)
                    If line.Length = 80 Then
                        If line.Substring(0, 4).Contains("ATOM") And Not (line.Substring(17, 3).Contains("DG") Or line.Substring(17, 3).Contains("DA") Or line.Substring(17, 3).Contains("DU") Or line.Substring(17, 3).Contains("DT") Or line.Substring(17, 3).Contains("DC") Or line.Substring(17, 3).Contains("DI")) Then
                            Dim substring As String = line.Substring(17, 3)
                            Select Case substring
                                Case "ALA"    'Alanine
                                    Acounter = Acounter + 1
                                    generalcounter = generalcounter + 1
                                Case "CYS"    'Cysteine
                                    Ccounter = Ccounter + 1
                                    generalcounter = generalcounter + 1
                                Case "ASP"    'Aspartate
                                    Dcounter = Dcounter + 1
                                    generalcounter = generalcounter + 1
                                Case "GLU"    'Glutamate
                                    Ecounter = Ecounter + 1
                                    generalcounter = generalcounter + 1
                                Case "PHE"    'Phenylalanine
                                    Fcounter = Fcounter + 1
                                    generalcounter = generalcounter + 1
                                Case "GLY"    'Glycine
                                    Gcounter = Gcounter + 1
                                    generalcounter = generalcounter + 1
                                Case "HIS"    'Histidine
                                    Hcounter = Hcounter + 1
                                    generalcounter = generalcounter + 1
                                Case "ILE"    'Isoleucine
                                    Icounter = Icounter + 1
                                    generalcounter = generalcounter + 1
                                Case "LYS"    'Lysine
                                    Kcounter = Kcounter + 1
                                    generalcounter = generalcounter + 1
                                Case "LEU"    'Leucine
                                    Lcounter = Lcounter + 1
                                    generalcounter = generalcounter + 1
                                Case "MET"    'Methionine
                                    Mcounter = Mcounter + 1
                                    generalcounter = generalcounter + 1
                                Case "ASN"    'Asparagine
                                    Ncounter = Ncounter + 1
                                    generalcounter = generalcounter + 1
                                Case "PRO"    'Proline
                                    Pcounter = Pcounter + 1
                                    generalcounter = generalcounter + 1
                                Case "GLN"    'Glutamine
                                    Qcounter = Qcounter + 1
                                    generalcounter = generalcounter + 1
                                Case "ARG"    'Arginine
                                    Rcounter = Rcounter + 1
                                    generalcounter = generalcounter + 1
                                Case "SER"    'Serine
                                    Scounter = Scounter + 1
                                    generalcounter = generalcounter + 1
                                Case "THR"    'Threonine
                                    Tcounter = Tcounter + 1
                                    generalcounter = generalcounter + 1
                                Case "VAL"    'Valine
                                    Vcounter = Vcounter + 1
                                    generalcounter = generalcounter + 1
                                Case "TRP"    'Tryptophan
                                    Wcounter = Wcounter + 1
                                    generalcounter = generalcounter + 1
                                Case "TYR"    'Tyrosine
                                    Ycounter = Ycounter + 1
                                    generalcounter = generalcounter + 1
                                Case "SEC"  'Selenocysteine
                                    'do nothing - skip
                                Case "GLX"  'GLU/GLN Ambiguous
                                    'do nothing - skip
                                Case "UNK"    'Unknown
                                    'do nothing - skip
                                Case "  U"    'Unknown
                                    'do nothing - skip
                                Case "  C"    'Unknown
                                    'do nothing - skip
                                Case "  G"    'Unknown
                                    'do nothing - skip
                                Case "  A"    'Unknown
                                    'do nothing - skip
                                Case "  T"    'Unknown
                                    'do nothing - skip
                                Case "  N"    'Unknown
                                    'do nothing - skip
                                Case Else
                                    Console.WriteLine(" unhandled character: " & f.ToString & " - " & substring & " - " & line)
                            End Select
                        End If
                    End If
                Loop
                filecounter = filecounter + 1
                Console.SetCursorPosition(0, Console.CursorTop)
                Console.Write(filecounter.ToString)
            Next


            ''Print Sequence Quantities
            Console.WriteLine(vbCrLf & vbCrLf & "--------- Sequence Results ---------")
            Console.WriteLine(vbCrLf & generalcounter & " sequence characters were processed." & vbCrLf)
            Console.WriteLine("1.  Alanine - Ala - A: " & Acounter & " characters - " & Math.Round(Acounter / generalcounter * 100, 2) & " %")
            Console.WriteLine("2.  Cysteine - Cys - C: " & Ccounter & " characters - " & Math.Round(Ccounter / generalcounter * 100, 2) & " %")
            Console.WriteLine("3.  Aspartate - Asp - D: " & Dcounter & " characters - " & Math.Round(Dcounter / generalcounter * 100, 2) & " %")
            Console.WriteLine("4.  Glutamate - Glu - E: " & Ecounter & " characters - " & Math.Round(Ecounter / generalcounter * 100, 2) & " %")
            Console.WriteLine("5.  Phenylalanine - Phe - F: " & Fcounter & " characters - " & Math.Round(Fcounter / generalcounter * 100, 2) & " %")
            Console.WriteLine("6.  Glycine - Gly - G: " & Gcounter & " characters - " & Math.Round(Gcounter / generalcounter * 100, 2) & " %")
            Console.WriteLine("7.  Histidine - His - H: " & Hcounter & " characters - " & Math.Round(Hcounter / generalcounter * 100, 2) & " %")
            Console.WriteLine("8.  Isoleucine - Ile - I: " & Icounter & " characters - " & Math.Round(Icounter / generalcounter * 100, 2) & " %")
            Console.WriteLine("9.  Lysine - Lys - K: " & Kcounter & " characters - " & Math.Round(Kcounter / generalcounter * 100, 2) & " %")
            Console.WriteLine("10. Leucine - Leu - L: " & Lcounter & " characters - " & Math.Round(Lcounter / generalcounter * 100, 2) & " %")
            Console.WriteLine("11. Methionine - Met - M: " & Mcounter & " characters - " & Math.Round(Mcounter / generalcounter * 100, 2) & " %")
            Console.WriteLine("12. Asparagine - Asn - N: " & Ncounter & " characters - " & Math.Round(Ncounter / generalcounter * 100, 2) & " %")
            Console.WriteLine("13. Proline - Pro - P: " & Pcounter & " characters - " & Math.Round(Pcounter / generalcounter * 100, 2) & " %")
            Console.WriteLine("14. Glutamine - Gln - Q: " & Qcounter & " characters - " & Math.Round(Qcounter / generalcounter * 100, 2) & " %")
            Console.WriteLine("15. Arginine - Arg - R: " & Rcounter & " characters - " & Math.Round(Rcounter / generalcounter * 100, 2) & " %")
            Console.WriteLine("16. Serine - Ser - S: " & Scounter & " characters - " & Math.Round(Scounter / generalcounter * 100, 2) & " %")
            Console.WriteLine("17. Threonine - Thr - T: " & Tcounter & " characters - " & Math.Round(Tcounter / generalcounter * 100, 2) & " %")
            Console.WriteLine("18. Valine - Val - V: " & Vcounter & " characters - " & Math.Round(Vcounter / generalcounter * 100, 2) & " %")
            Console.WriteLine("19. Tryptophan - Trp - W: " & Wcounter & " characters - " & Math.Round(Wcounter / generalcounter * 100, 2) & " %")
            Console.WriteLine("20. Tyrosine - Tyr - Y: " & Ycounter & " characters - " & Math.Round(Ycounter / generalcounter * 100, 2) & " %" & vbCrLf)


            'create the 1st worksheet
            ws = pck.Workbook.Worksheets.Add("Amino Acids Ratio")
            ws.Cells.AutoFitColumns()
            ws.Cells.Style.Fill.PatternType = Style.ExcelFillStyle.Solid
            ws.Cells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White)

            ws.Cells("A1").Value = "Date: " & Now.ToShortDateString
            ws.Cells("A2").Value = "Report: Amino Acids Ratio"
            ws.Cells("A3").Value = "Processed Amino Acids: " + generalcounter.ToString
            ws.Cells("A1:A3").Style.Font.Bold = True

            'ws.Cells.AutoFitColumns(25)
            ws.Column(1).Width = 35
            ws.Column(2).Width = 20
            ws.Column(3).Width = 20

            'header
            ws.Cells(5, 1, 5, 3).Style.Font.Bold = True
            ws.Cells(5, 1, 5, 3).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
            ws.Cells(5, 1, 5, 3).Style.WrapText() = True
            ws.Cells(5, 1, 5, 3).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center
            ws.Cells(5, 1, 5, 3).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
            ws.Cells(5, 1, 5, 3).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 192, 192, 192))

            'borders
            ws.Cells(5, 1, 25, 3).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
            ws.Cells(5, 1, 25, 3).Style.Border.Bottom.Style = Style.ExcelBorderStyle.Thin
            ws.Cells(5, 1, 25, 3).Style.Border.Top.Style = Style.ExcelBorderStyle.Thin
            ws.Cells(5, 1, 25, 3).Style.Border.Left.Style = Style.ExcelBorderStyle.Thin
            ws.Cells(5, 1, 25, 3).Style.Border.Right.Style = Style.ExcelBorderStyle.Thin
            ws.Cells(5, 1, 25, 3).Style.Border.BorderAround(Style.ExcelBorderStyle.Thick)

            ws.Cells("A5").Value = "Amino Acid"
            ws.Cells("B5").Value = "Total Number of A/A"
            ws.Cells("C5").Value = "Percentage (%)"


            ws.Cells("A6").Value = "Alanine - Ala - A"
            ws.Cells("B6").Value = Acounter.ToString
            ws.Cells("C6").Value = Math.Round(Acounter / generalcounter * 100, 2).ToString

            ws.Cells("A7").Value = "Cysteine - Cys - C"
            ws.Cells("B7").Value = Ccounter.ToString
            ws.Cells("C7").Value = Math.Round(Ccounter / generalcounter * 100, 2).ToString

            ws.Cells("A8").Value = "Aspartate - Asp - D"
            ws.Cells("B8").Value = Dcounter.ToString
            ws.Cells("C8").Value = Math.Round(Dcounter / generalcounter * 100, 2).ToString

            ws.Cells("A9").Value = "Glutamate - Glu - E"
            ws.Cells("B9").Value = Ecounter.ToString
            ws.Cells("C9").Value = Math.Round(Ecounter / generalcounter * 100, 2).ToString

            ws.Cells("A10").Value = "Phenylalanine - Phe - F"
            ws.Cells("B10").Value = Fcounter.ToString
            ws.Cells("C10").Value = Math.Round(Fcounter / generalcounter * 100, 2).ToString

            ws.Cells("A11").Value = "Glycine - Gly - G"
            ws.Cells("B11").Value = Gcounter.ToString
            ws.Cells("C11").Value = Math.Round(Gcounter / generalcounter * 100, 2).ToString

            ws.Cells("A12").Value = "Histidine - His - H"
            ws.Cells("B12").Value = Hcounter.ToString
            ws.Cells("C12").Value = Math.Round(Hcounter / generalcounter * 100, 2).ToString

            ws.Cells("A13").Value = "Isoleucine - Ile - I"
            ws.Cells("B13").Value = Icounter.ToString
            ws.Cells("C13").Value = Math.Round(Icounter / generalcounter * 100, 2).ToString

            ws.Cells("A14").Value = "Lysine - Lys - K"
            ws.Cells("B14").Value = Kcounter.ToString
            ws.Cells("C14").Value = Math.Round(Kcounter / generalcounter * 100, 2).ToString

            ws.Cells("A15").Value = "Leucine - Leu - L"
            ws.Cells("B15").Value = Lcounter.ToString
            ws.Cells("C15").Value = Math.Round(Lcounter / generalcounter * 100, 2).ToString

            ws.Cells("A16").Value = "Methionine - Met - M"
            ws.Cells("B16").Value = Mcounter.ToString
            ws.Cells("C16").Value = Math.Round(Mcounter / generalcounter * 100, 2).ToString

            ws.Cells("A17").Value = "Asparagine - Asn - N"
            ws.Cells("B17").Value = Ncounter.ToString
            ws.Cells("C17").Value = Math.Round(Ncounter / generalcounter * 100, 2).ToString

            ws.Cells("A18").Value = "Proline - Pro - P"
            ws.Cells("B18").Value = Pcounter.ToString
            ws.Cells("C18").Value = Math.Round(Pcounter / generalcounter * 100, 2).ToString

            ws.Cells("A19").Value = "Glutamine - Gln - Q"
            ws.Cells("B19").Value = Qcounter.ToString
            ws.Cells("C19").Value = Math.Round(Qcounter / generalcounter * 100, 2).ToString

            ws.Cells("A20").Value = "Arginine - Arg - R"
            ws.Cells("B20").Value = Rcounter.ToString
            ws.Cells("C20").Value = Math.Round(Rcounter / generalcounter * 100, 2).ToString

            ws.Cells("A21").Value = "Serine - Ser - S"
            ws.Cells("B21").Value = Scounter.ToString
            ws.Cells("C21").Value = Math.Round(Scounter / generalcounter * 100, 2).ToString

            ws.Cells("A22").Value = "Threonine - Thr - T"
            ws.Cells("B22").Value = Tcounter.ToString
            ws.Cells("C22").Value = Math.Round(Tcounter / generalcounter * 100, 2).ToString

            ws.Cells("A23").Value = "Valine - Val - V"
            ws.Cells("B23").Value = Vcounter.ToString
            ws.Cells("C23").Value = Math.Round(Vcounter / generalcounter * 100, 2).ToString

            ws.Cells("A24").Value = "Tryptophan - Trp - W"
            ws.Cells("B24").Value = Wcounter.ToString
            ws.Cells("C24").Value = Math.Round(Wcounter / generalcounter * 100, 2).ToString

            ws.Cells("A25").Value = "Tyrosine - Tyr - Y"
            ws.Cells("B25").Value = Ycounter.ToString
            ws.Cells("C25").Value = Math.Round(Ycounter / generalcounter * 100, 2).ToString

            Dim P As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)
            Dim strPath As String = New Uri(P).LocalPath

            pck.SaveAs(New FileInfo(strPath & "\Amino_Acids_Ratio_Report.xlsx"))
            Console.WriteLine(vbCrLf & "A report file was generated successfully under the following path. " & vbCrLf & strPath)

        Catch e As Exception

            Console.WriteLine("An error occured.")
            Console.WriteLine("The error message is: " & e.Message)
            Console.WriteLine(vbCrLf & "Press enter to exit...")
            Console.ReadLine()
            Environment.Exit(0)
        End Try
        stpw.Stop()
        Console.WriteLine(vbCrLf & vbCrLf & "The application was terminated.")
        Console.WriteLine("Elapsed time: " & stpw.Elapsed.ToString)
        Console.WriteLine("Press enter to exit...")
        Console.ReadLine()
        Environment.Exit(0)
    End Sub

End Module
