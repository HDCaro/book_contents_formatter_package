Imports System
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Collections.Generic
Imports Microsoft.Office.Interop.Word
Imports WordRange = Microsoft.Office.Interop.Word.Range

Module FrontMatterBuilder

    ' -----------------------------------
    ' WORD CONSTANTS
    ' -----------------------------------
    Const wdRestartContinuous As Integer = 0
    Const wdRestartPage As Integer = 2
    Const wdRestartSection As Integer = 1

    Const wdPageNumberStyleLowercaseRoman As Integer = 14
    Const wdPageNumberStyleUppercaseRoman As Integer = 13

    Const wdAlignParagraphCenter As Integer = 1

    Const wdFieldPage As Integer = 33

    Const wdSectionBreakNextPage As Integer = 2
    Const wdPageBreak As Integer = 7

    Const wdHeaderFooterPrimary As Integer = 1

    ' -----------------------------------
    ' FIND PROJECT ROOT
    ' -----------------------------------
    Function FindProjectRoot() As String
        Dim current As String = Path.GetFullPath(GetType(FrontMatterBuilder).Assembly.Location)
        Dim dir As String = Path.GetDirectoryName(current)

        Dim candidate As String = dir
        Do While Not String.IsNullOrEmpty(candidate)
            If Directory.Exists(Path.Combine(candidate, "src")) AndAlso
               Directory.Exists(Path.Combine(candidate, "data")) Then
                Return candidate
            End If
            Dim parent As String = Path.GetDirectoryName(candidate)
            If parent = candidate Then Exit Do
            candidate = parent
        Loop

        Throw New Exception("Project root not found (expected folders: src and data)")
    End Function

    ' -----------------------------------
    ' FIND FILE BY NAME (SEARCH RECURSIVELY)
    ' -----------------------------------
    Function FindFileByName(projectRoot As String, filename As String) As List(Of String)
        Dim matches As New List(Of String)
        Dim target As String = filename.ToLower()

        For Each filePath As String In Directory.EnumerateFiles(projectRoot, "*", SearchOption.AllDirectories)
            If Path.GetFileName(filePath).ToLower() = target Then
                matches.Add(filePath)
            End If
        Next

        ' Sort by depth (shallower first) then alphabetically
        matches.Sort(Function(a, b)
                         Dim depthA = a.Split(Path.DirectorySeparatorChar).Length
                         Dim depthB = b.Split(Path.DirectorySeparatorChar).Length
                         Dim cmp = depthA.CompareTo(depthB)
                         If cmp <> 0 Then Return cmp
                         Return String.Compare(a, b, StringComparison.OrdinalIgnoreCase)
                     End Function)

        Return matches
    End Function

    ' -----------------------------------
    ' RESOLVE EXISTING PATH
    ' -----------------------------------
    Function ResolveExistingPath(projectRoot As String, label As String,
                                  preferredRelativePaths As IEnumerable(Of String),
                                  filename As String) As String
        For Each rel As String In preferredRelativePaths
            Dim candidate As String = Path.Combine(projectRoot, rel)
            If File.Exists(candidate) Then
                Console.WriteLine($"   [OK] Resolved {label}: {MakeRelative(projectRoot, candidate)}")
                Return candidate
            End If
        Next

        Dim matches = FindFileByName(projectRoot, filename)
        If matches.Count > 0 Then
            Dim resolved = matches(0)
            Console.WriteLine($"   [OK] Auto-discovered {label}: {MakeRelative(projectRoot, resolved)}")
            Return resolved
        End If

        Dim searchList = String.Join(Environment.NewLine & "      - ",
                                     preferredRelativePaths)
        Throw New FileNotFoundException(
            $"Could not find {label} ('{filename}') under project root: {projectRoot}" &
            Environment.NewLine & $"Checked preferred paths:{Environment.NewLine}      - {searchList}")
    End Function

    Function MakeRelative(root As String, fullPath As String) As String
        If fullPath.StartsWith(root, StringComparison.OrdinalIgnoreCase) Then
            Return fullPath.Substring(root.Length).TrimStart(Path.DirectorySeparatorChar)
        End If
        Return fullPath
    End Function

    ' -----------------------------------
    ' CLEAN TEXT (REMOVE CONTROL CHARS)
    ' -----------------------------------
    Function CleanText(text As String) As String
        If String.IsNullOrEmpty(text) Then Return ""

        Dim sb As New StringBuilder()
        For Each c As Char In text
            If c >= " "c OrElse c = vbCr(0) OrElse c = vbLf(0) OrElse c = vbTab(0) Then
                sb.Append(c)
            End If
        Next

        Return sb.ToString().TrimEnd("/"c, "\"c).Trim()
    End Function

    ' -----------------------------------
    ' CLEAN TITLE TEXT (REMOVE ALL LINE BREAKS)
    ' -----------------------------------
    Function CleanTitleText(text As String) As String
        If String.IsNullOrEmpty(text) Then Return ""

        Dim cleaned = text.Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ")

        ' Normalize whitespace
        cleaned = Regex.Replace(cleaned, "\s+", " ")

        ' Remove control characters except spaces
        Dim sb As New StringBuilder()
        For Each c As Char In cleaned
            If c >= " "c Then sb.Append(c)
        Next

        Return sb.ToString().TrimEnd("/"c, "\"c).Trim()
    End Function

    ' -----------------------------------
    ' GET WORD APPLICATION
    ' -----------------------------------
    Function GetWord() As Application
        Dim word As New Application()
        word.Visible = False
        word.DisplayAlerts = WdAlertLevel.wdAlertsNone
        Return word
    End Function

    ' -----------------------------------
    ' EXTRACT HEADINGS
    ' -----------------------------------
    Function ExtractHeadings(docPath As String) As List(Of Dictionary(Of String, Object))
        Console.WriteLine(vbCrLf & "Extracting headings with consecutive Heading 2 collection")

        Dim word As Application = GetWord()
        Dim doc As Document = word.Documents.Open(Path.GetFullPath(docPath))
        doc.Repaginate()

        Dim headings As New List(Of Dictionary(Of String, Object))
        Dim total As Integer = doc.Paragraphs.Count

        Dim i As Integer = 1
        Do While i <= total
            Try
                Dim para As Paragraph = doc.Paragraphs(i)
                Dim text As String = CleanText(para.Range.Text)
                Dim style As String = para.Style.NameLocal.ToLower()

                ' Look for Heading 1 (Chapter numbers)
                If style.Contains("heading 1") AndAlso Not String.IsNullOrEmpty(text) Then
                    Console.WriteLine($"{vbCrLf}   Found Heading 1: '{text}'")

                    ' ---- Chapter with Heading 2 title(s) ----
                    If text.ToLower().StartsWith("chapter") Then
                        Console.WriteLine($"   Processing chapter: '{text}'")

                        Dim chapter As String = text
                        Dim chapterTitleParts As New List(Of String)
                        Dim j As Integer = i + 1

                        Console.WriteLine($"   Looking for consecutive Heading 2 titles starting from paragraph {j}...")

                        Do While j <= total
                            Try
                                Dim p As Paragraph = doc.Paragraphs(j)
                                Dim pStyle As String = p.Style.NameLocal.ToLower()
                                Dim rawText As String = p.Range.Text
                                Dim cleanedText As String = CleanTitleText(rawText)

                                Console.WriteLine($"   Paragraph {j}: '{cleanedText}' (style: {pStyle})")

                                If String.IsNullOrWhiteSpace(rawText) Then
                                    Console.WriteLine($"   Empty paragraph {j}, skipping")
                                    j += 1
                                    Continue Do
                                End If

                                If pStyle.Contains("heading 2") Then
                                    chapterTitleParts.Add(cleanedText)
                                    Console.WriteLine($"   Collected Heading 2 part: '{cleanedText}'")
                                    j += 1
                                    Continue Do
                                ElseIf pStyle.Contains("heading 1") Then
                                    Console.WriteLine("   Hit another Heading 1, stopping search")
                                    Exit Do
                                Else
                                    Console.WriteLine("   Hit non-heading style, stopping title collection")
                                    Exit Do
                                End If

                            Catch ex As Exception
                                Console.WriteLine($"   Warning: Error processing paragraph {j}: {ex.Message}")
                                Exit Do
                            End Try
                        Loop

                        Dim full As String
                        If chapterTitleParts.Count > 0 Then
                            Dim completeTitle = String.Join(" ", chapterTitleParts)
                            full = $"{chapter}: {completeTitle}"
                            Console.WriteLine($"   Complete chapter title: '{full}'")
                        Else
                            full = chapter
                            Console.WriteLine($"   No Heading 2 found, using chapter only: '{full}'")
                        End If

                        full = CleanTitleText(full)
                        Dim page As Object = para.Range.Information(WdInformation.wdActiveEndPageNumber)

                        Dim entry As New Dictionary(Of String, Object)
                        entry("text") = full
                        entry("page") = page
                        headings.Add(entry)
                        Console.WriteLine($"   Added to TOC: '{full}' -> page {page}")

                        i = j
                        Continue Do

                    Else
                        ' ---- Single-line heading ----
                        Dim page As Object = para.Range.Information(WdInformation.wdActiveEndPageNumber)
                        Dim cleanHeading = CleanTitleText(text)
                        Dim entry As New Dictionary(Of String, Object)
                        entry("text") = cleanHeading
                        entry("page") = page
                        headings.Add(entry)
                        Console.WriteLine($"   Single-line heading: '{cleanHeading}' -> page {page}")
                    End If

                ElseIf style.Contains("heading 2") AndAlso Not String.IsNullOrEmpty(text) Then
                    Dim page As Object = para.Range.Information(WdInformation.wdActiveEndPageNumber)
                    Dim cleanHeading = CleanTitleText(para.Range.Text)
                    Dim entry As New Dictionary(Of String, Object)
                    entry("text") = cleanHeading
                    entry("page") = page
                    headings.Add(entry)
                    Console.WriteLine($"   Standalone Heading 2: '{cleanHeading}' -> page {page}")
                End If

            Catch ex As Exception
                Console.WriteLine($"   Warning: Error processing paragraph {i}: {ex.Message}")
            End Try

            i += 1
        Loop

        doc.Close(False)
        word.Quit()

        Console.WriteLine($"{vbCrLf}   Total headings extracted: {headings.Count}")
        For idx As Integer = 0 To headings.Count - 1
            Console.WriteLine($"   {idx + 1}. {headings(idx)("text")} -> page {headings(idx)("page")}")
        Next

        Return headings
    End Function

    ' -----------------------------------
    ' BUILD TOC DOC
    ' -----------------------------------
    Function BuildTocDoc(headings As List(Of Dictionary(Of String, Object)), tocPath As String) As String
        Console.WriteLine(vbCrLf & "Building TOC with guaranteed single-line titles")

        Dim word As Application = GetWord()
        Dim doc As Document = word.Documents.Add()

        Try
            ' Title paragraph
            Dim titlePara As Paragraph = doc.Content.Paragraphs.Add()
            titlePara.Range.Text = "CONTENTS"
            titlePara.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
            titlePara.Range.Font.Name = "Georgia"
            titlePara.Range.Font.Size = 20
            titlePara.Range.Font.Bold = True
            titlePara.Range.InsertParagraphAfter()

            ' Blank line
            Dim blankPara As Paragraph = doc.Content.Paragraphs.Add()
            blankPara.Range.Text = ""
            blankPara.Range.InsertParagraphAfter()

            For Each h In headings
                Dim para As Paragraph = doc.Content.Paragraphs.Add()

                ' Hanging indent (0.5 inch = 720 twips)
                para.Format.LeftIndent = 720
                para.Format.FirstLineIndent = -720

                ' Tab stop at 6 inches with dot leader (right-aligned)
                ' 6 inches = 8640 twips
                para.Format.TabStops.Add(8640,
                    WdTabAlignment.wdAlignTabRight,
                    WdTabLeader.wdTabLeaderDots)

                Dim titleText As String = CStr(h("text"))
                ' Final safety: remove line breaks
                titleText = titleText.Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ")
                titleText = Regex.Replace(titleText, "\s+", " ").Trim()

                Console.WriteLine($"   TOC entry: '{titleText}' -> {h("page")}")

                ' Write text + tab + page number
                para.Range.Text = $"{titleText}{vbTab}{h("page")}"
                para.Range.Font.Name = "Georgia"
                para.Range.Font.Size = 12

                para.Range.InsertParagraphAfter()
            Next

            doc.SaveAs2(Path.GetFullPath(tocPath))
            doc.Close(False)
            word.Quit()

            Console.WriteLine("   TOC created with guaranteed single-line entries")

        Catch ex As Exception
            Try : word.Quit() : Catch : End Try
            Throw
        End Try

        Return tocPath
    End Function

    ' -----------------------------------
    ' APPLY ROMAN PAGINATION
    ' -----------------------------------
    Sub ApplyRomanPagination(doc As Document)
        Console.WriteLine(vbCrLf & "Applying Roman pagination (FIXED FORMAT)")

        Try
            Dim sections As Sections = doc.Sections
            Console.WriteLine($"   Total sections: {sections.Count}")

            If sections.Count < 2 Then
                Console.WriteLine("   Warning: Not enough sections")
                Return
            End If

            ' SECTION 1 - No page numbers (title page)
            Dim sec1 As Section = sections.Item(1)
            Try
                sec1.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Delete()
                sec1.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Delete()
                sec1.PageSetup.DifferentFirstPageHeaderFooter = True
                Console.WriteLine("   Section 1: no numbering")
            Catch ex As Exception
                Console.WriteLine($"   Warning: Section 1 cleanup error: {ex.Message}")
            End Try

            ' SECTION 2 - Roman numbers
            Dim sec2 As Section = sections.Item(2)

            Try
                sec2.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = False
                sec2.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = False
                Console.WriteLine("   Section 2 unlinked")
            Catch ex As Exception
                Console.WriteLine($"   Warning: Unlink failed: {ex.Message}")
            End Try

            Try
                Console.WriteLine("   Setting page restart properties...")

                Dim footer As HeaderFooter = sec2.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                Dim pageNums As PageNumbers = footer.PageNumbers

                Console.WriteLine($"   Removing {pageNums.Count} existing page numbers...")
                Do While pageNums.Count > 0
                    pageNums.Item(1).Delete()
                Loop

                Console.WriteLine("   Adding Roman page numbers...")
                pageNums.Add(WdPageNumberAlignment.wdAlignPageNumberCenter)
                pageNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleLowercaseRoman
                pageNums.RestartNumberingAtSection = True
                pageNums.StartingNumber = 1
                Console.WriteLine("   Page restart: True, Start: 1")

                Console.WriteLine($"   Roman format applied: NumberStyle = {wdPageNumberStyleLowercaseRoman}")
                Console.WriteLine("   Roman numbering configured (i, ii, iii...)")

            Catch ex As Exception
                Console.WriteLine($"   ERROR: Page number setup failed: {ex.Message}")

                ' Fallback: insert PAGE field directly
                Try
                    Console.WriteLine("   Trying fallback field method...")
                    Dim footer As HeaderFooter = sec2.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                    footer.Range.Delete()

                    Dim footerRange As WordRange = footer.Range
                    Dim field As Field = footerRange.Fields.Add(
                        Range:=footerRange,
                        Type:=WdFieldType.wdFieldPage,
                        PreserveFormatting:=False)
                    field.Code.Text = "PAGE \* ROMAN \* LOWER"
                    field.Update()

                    footer.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
                    Console.WriteLine("   Fallback Roman field inserted")

                Catch ex2 As Exception
                    Console.WriteLine($"   ERROR: Fallback method also failed: {ex2.Message}")
                End Try
            End Try

            ' Force document update
            Try
                Console.WriteLine("   Forcing document updates...")
                doc.ActiveWindow.View.ShowFieldCodes = False
                doc.Repaginate()
                doc.Fields.Update()
                doc.Repaginate()
                Console.WriteLine("   Document refreshed")
            Catch ex As Exception
                Console.WriteLine($"   Warning: Refresh error: {ex.Message}")
            End Try

            ' Verification
            Try
                Console.WriteLine("   Verifying Roman format...")
                Dim sec2Footer As HeaderFooter = sec2.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                Dim sec2PageNums As PageNumbers = sec2Footer.PageNumbers
                If sec2PageNums.Count > 0 Then
                    Dim style As Integer = CInt(sec2PageNums.NumberStyle)
                    Console.WriteLine($"   Current NumberStyle: {style} (should be {wdPageNumberStyleLowercaseRoman})")
                    If style = wdPageNumberStyleLowercaseRoman Then
                        Console.WriteLine("   Roman format verified!")
                    Else
                        Console.WriteLine("   Warning: Roman format not applied correctly")
                    End If
                Else
                    Console.WriteLine("   Warning: No page numbers found for verification")
                End If
            Catch ex As Exception
                Console.WriteLine($"   Warning: Verification failed: {ex.Message}")
            End Try

        Catch ex As Exception
            Console.WriteLine($"   ERROR: Pagination error: {ex.Message}")
        End Try
    End Sub

    ' -----------------------------------
    ' ASSEMBLE FINAL DOCUMENT
    ' -----------------------------------
    Sub AssembleFinal(titleDoc As String, copyrightDoc As String, tocDoc As String, output As String)
        Console.WriteLine(vbCrLf & "Assembling final document with Roman pagination fix")

        Dim word As Application = GetWord()
        Dim doc As Document = word.Documents.Add()

        Try
            ' Title page (Section 1) - no page numbering
            If File.Exists(titleDoc) Then
                Console.WriteLine("   Inserting title page (Section 1)...")
                Dim r As WordRange = doc.Content
                r.Collapse(WdCollapseDirection.wdCollapseEnd)
                r.InsertFile(Path.GetFullPath(titleDoc))

                r = doc.Content
                r.Collapse(WdCollapseDirection.wdCollapseEnd)
                r.InsertBreak(WdBreakType.wdSectionBreakNextPage)
                Console.WriteLine("   Title page + section break inserted")
            End If

            ' Copyright page (Section 2, page "i")
            If File.Exists(copyrightDoc) Then
                Console.WriteLine("   Inserting copyright page (will be Roman 'i')...")
                Dim r As WordRange = doc.Content
                r.Collapse(WdCollapseDirection.wdCollapseEnd)
                r.InsertFile(Path.GetFullPath(copyrightDoc))

                r = doc.Content
                r.Collapse(WdCollapseDirection.wdCollapseEnd)
                r.InsertBreak(WdBreakType.wdPageBreak)
                Console.WriteLine("   Copyright page inserted")
            End If

            ' TOC (Section 2, pages "ii", "iii", ...)
            Console.WriteLine("   Inserting TOC (will be Roman 'ii', 'iii'...)...")
            Dim rToc As WordRange = doc.Content
            rToc.Collapse(WdCollapseDirection.wdCollapseEnd)
            rToc.InsertFile(Path.GetFullPath(tocDoc))
            Console.WriteLine("   TOC inserted")

            ' Apply Roman pagination
            ApplyRomanPagination(doc)

            ' Final processing
            Console.WriteLine("   Final document processing...")
            Try
                doc.ActiveWindow.View.ShowFieldCodes = False
                doc.Repaginate()
                doc.Fields.Update()
                doc.Repaginate()
            Catch ex As Exception
                Console.WriteLine($"   Warning: Final processing warning: {ex.Message}")
            End Try

            ' Save
            Console.WriteLine("   Saving document...")
            doc.SaveAs2(Path.GetFullPath(output))
            doc.Close(False)
            word.Quit()

            Console.WriteLine(vbCrLf & "DOCUMENT COMPLETED!")
            Console.WriteLine($"File: {output}")
            Console.WriteLine(vbCrLf & "Expected result:")
            Console.WriteLine("   - Title page: No number")
            Console.WriteLine("   - Copyright page: Roman 'i'")
            Console.WriteLine("   - TOC pages: Roman 'ii', 'iii', 'iv'...")
            Console.WriteLine("   - ALL chapter titles on single lines")
            Console.WriteLine(vbCrLf & "Roman format fix applied:")
            Console.WriteLine("   - NumberStyle set on PageNumbers collection (not PageSetup)")
            Console.WriteLine("   - Fallback field method if PageNumbers fails")

        Catch ex As Exception
            Try : word.Quit() : Catch : End Try
            Console.WriteLine($"ERROR: Assembly error: {ex.Message}")
        End Try
    End Sub

    ' -----------------------------------
    ' ENTRY POINT (call from Program.Main)
    ' -----------------------------------
    Sub BuildFrontMatter()
        Dim projectRoot As String = FindProjectRoot()

        Dim bookFilename As String = "HITS AND HAPPINESS FINAL 2 Format MOM Discog.docx"
        Dim titleFilename As String = "HH Title.docx"
        Dim copyrightFilename As String = "HH Copyright page.docx"

        Dim bookFile As String = ResolveExistingPath(
            projectRoot,
            "book file",
            {
                Path.Combine("data", "index", "input", "book", bookFilename),
                Path.Combine("data", "index", "input", bookFilename),
                Path.Combine("release", "v1", bookFilename)
            },
            bookFilename)

        Dim titleFile As String = ResolveExistingPath(
            projectRoot,
            "title file",
            {
                titleFilename,
                Path.Combine("data", "front_mater", "input", titleFilename),
                Path.Combine("release", "v1", titleFilename)
            },
            titleFilename)

        Dim copyrightFile As String = ResolveExistingPath(
            projectRoot,
            "copyright file",
            {
                copyrightFilename,
                Path.Combine("data", "front_mater", "input", copyrightFilename),
                Path.Combine("release", "v1", copyrightFilename)
            },
            copyrightFilename)

        Dim tocTempDir As String = Path.Combine(projectRoot, "data", "front_mater", "output")
        Directory.CreateDirectory(tocTempDir)
        Dim tocTemp As String = Path.Combine(tocTempDir, "temp_toc.docx")

        Dim outputDir As String = Path.Combine(projectRoot, "release", "v1")
        Directory.CreateDirectory(outputDir)
        Dim bookStem As String = Path.GetFileNameWithoutExtension(bookFile)
        Dim outputFile As String = Path.Combine(outputDir, $"{bookStem} TOC.docx")

        Console.WriteLine($"{vbCrLf}Project root: {projectRoot}")
        Console.WriteLine($"Book: {MakeRelative(projectRoot, bookFile)}")
        Console.WriteLine($"Title: {MakeRelative(projectRoot, titleFile)}")
        Console.WriteLine($"Copyright: {MakeRelative(projectRoot, copyrightFile)}")
        Console.WriteLine($"Temp TOC: {MakeRelative(projectRoot, tocTemp)}")
        Console.WriteLine($"Output: {MakeRelative(projectRoot, outputFile)}")

        Dim headings = ExtractHeadings(bookFile)

        If headings.Count > 0 Then
            BuildTocDoc(headings, tocTemp)
            AssembleFinal(titleFile, copyrightFile, tocTemp, outputFile)

            If File.Exists(tocTemp) Then
                File.Delete(tocTemp)
            End If
        End If
    End Sub

End Module
