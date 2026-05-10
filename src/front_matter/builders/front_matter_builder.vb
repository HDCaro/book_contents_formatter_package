Imports System
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Text.Json
Imports System.Collections.Generic
Imports Microsoft.Office.Interop.Word
Imports BookAutomationCore
Imports WordRange = Microsoft.Office.Interop.Word.Range

Module FrontMatterBuilder

    Private Class BuilderConfig
        Public Property ConfigPath As String
        Public Property TitleFile As String
        Public Property CopyrightFile As String
        Public Property BookBodyFile As String
        Public Property OutputDir As String
        Public Property OutputFilename As String
        Public Property MetadataFilename As String
        Public Property TempDir As String
        Public Property DeleteTempToc As Boolean
        Public Property ApplyBookLayout As Boolean
        Public Property PageNumberingStyle As String
        Public Property PageWidthTwipsOverride As Integer?
        Public Property PageHeightTwipsOverride As Integer?
    End Class

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

    Function ResolveConfigPath(projectRoot As String, configuredPath As String) As String
        If String.IsNullOrWhiteSpace(configuredPath) Then Return configuredPath
        If Path.IsPathRooted(configuredPath) Then Return Path.GetFullPath(configuredPath)
        Return Path.GetFullPath(Path.Combine(projectRoot, configuredPath))
    End Function

    Function TryGetNullableInt(options As JsonElement, propertyName As String) As Integer?
        Dim prop As JsonElement
        If options.ValueKind = JsonValueKind.Object AndAlso options.TryGetProperty(propertyName, prop) Then
            If prop.ValueKind = JsonValueKind.Number Then
                Dim value As Integer
                If prop.TryGetInt32(value) Then
                    Return value
                End If
            End If
        End If
        Return Nothing
    End Function

    Function LoadBuilderConfig(projectRoot As String) As BuilderConfig
        Dim configPath As String = Path.Combine(projectRoot, "src", "front_matter", "builders", "front_matter_builder.config.json")
        If Not File.Exists(configPath) Then
            Throw New FileNotFoundException("Configuration file not found", configPath)
        End If

        Dim json As String = File.ReadAllText(configPath, Encoding.UTF8)
        Using doc As JsonDocument = JsonDocument.Parse(json)
            Dim root = doc.RootElement

            Dim inputs As JsonElement
            If Not root.TryGetProperty("inputs", inputs) Then
                Throw New Exception("Invalid front_matter_builder.config.json: missing 'inputs'")
            End If

            Dim outputs As JsonElement
            If Not root.TryGetProperty("outputs", outputs) Then
                Throw New Exception("Invalid front_matter_builder.config.json: missing 'outputs'")
            End If

            Dim options As JsonElement
            root.TryGetProperty("options", options)

            Dim titleConfigured As String = inputs.GetProperty("title_file").GetString()
            Dim copyrightConfigured As String = inputs.GetProperty("copyright_file").GetString()
            Dim bookConfigured As String = inputs.GetProperty("book_body_file").GetString()

            Dim autoDiscover As Boolean = True
            Dim autoDiscoverProp As JsonElement
            If inputs.TryGetProperty("auto_discover_missing_inputs", autoDiscoverProp) AndAlso autoDiscoverProp.ValueKind = JsonValueKind.False Then
                autoDiscover = False
            End If

            Dim titleFile As String = ResolveConfigPath(projectRoot, titleConfigured)
            If Not File.Exists(titleFile) AndAlso autoDiscover Then
                Dim matches = FindFileByName(projectRoot, Path.GetFileName(titleConfigured))
                If matches.Count > 0 Then titleFile = matches(0)
            End If

            Dim copyrightFile As String = ResolveConfigPath(projectRoot, copyrightConfigured)
            If Not File.Exists(copyrightFile) AndAlso autoDiscover Then
                Dim matches = FindFileByName(projectRoot, Path.GetFileName(copyrightConfigured))
                If matches.Count > 0 Then copyrightFile = matches(0)
            End If

            Dim bookBodyFile As String = ResolveConfigPath(projectRoot, bookConfigured)
            If Not File.Exists(bookBodyFile) AndAlso autoDiscover Then
                Dim matches = FindFileByName(projectRoot, Path.GetFileName(bookConfigured))
                If matches.Count > 0 Then bookBodyFile = matches(0)
            End If

            If Not File.Exists(titleFile) Then Throw New FileNotFoundException("title_file not found", titleFile)
            If Not File.Exists(copyrightFile) Then Throw New FileNotFoundException("copyright_file not found", copyrightFile)
            If Not File.Exists(bookBodyFile) Then Throw New FileNotFoundException("book_body_file not found", bookBodyFile)

            Dim outputDir As String = ResolveConfigPath(projectRoot, outputs.GetProperty("output_dir").GetString())
            Dim tempDir As String = ResolveConfigPath(projectRoot, outputs.GetProperty("temp_dir").GetString())
            Dim outputFilename As String = outputs.GetProperty("filename").GetString()

            Dim metadataFilename As String = "front_matter.config.json"
            Dim metadataProp As JsonElement
            If outputs.TryGetProperty("metadata_filename", metadataProp) AndAlso metadataProp.ValueKind = JsonValueKind.String Then
                metadataFilename = metadataProp.GetString()
            End If

            Dim deleteTempToc As Boolean = True
            Dim applyBookLayout As Boolean = True
            Dim pageNumberingStyle As String = "roman_lowercase"
            Dim pageWidthTwipsOverride As Integer? = Nothing
            Dim pageHeightTwipsOverride As Integer? = Nothing

            If options.ValueKind = JsonValueKind.Object Then
                Dim deleteProp As JsonElement
                If options.TryGetProperty("delete_temp_toc", deleteProp) Then
                    deleteTempToc = deleteProp.ValueKind <> JsonValueKind.False
                End If

                Dim applyLayoutProp As JsonElement
                If options.TryGetProperty("apply_book_layout", applyLayoutProp) Then
                    applyBookLayout = applyLayoutProp.ValueKind <> JsonValueKind.False
                End If

                Dim styleProp As JsonElement
                If options.TryGetProperty("page_numbering_style", styleProp) AndAlso styleProp.ValueKind = JsonValueKind.String Then
                    pageNumberingStyle = styleProp.GetString()
                End If

                pageWidthTwipsOverride = TryGetNullableInt(options, "page_width_twips_override")
                pageHeightTwipsOverride = TryGetNullableInt(options, "page_height_twips_override")
            End If

            Return New BuilderConfig With {
                .ConfigPath = configPath,
                .TitleFile = titleFile,
                .CopyrightFile = copyrightFile,
                .BookBodyFile = bookBodyFile,
                .OutputDir = outputDir,
                .OutputFilename = outputFilename,
                .MetadataFilename = metadataFilename,
                .TempDir = tempDir,
                .DeleteTempToc = deleteTempToc,
                .ApplyBookLayout = applyBookLayout,
                .PageNumberingStyle = pageNumberingStyle,
                .PageWidthTwipsOverride = pageWidthTwipsOverride,
                .PageHeightTwipsOverride = pageHeightTwipsOverride
            }
        End Using
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
                Dim outlineLevel As WdOutlineLevel = para.OutlineLevel
                Dim isHeading1 As Boolean = (outlineLevel = WdOutlineLevel.wdOutlineLevel1) OrElse style.Contains("heading 1")
                Dim isHeading2 As Boolean = (outlineLevel = WdOutlineLevel.wdOutlineLevel2) OrElse style.Contains("heading 2")

                ' Look for Heading 1 (Chapter numbers)
                If isHeading1 AndAlso Not String.IsNullOrEmpty(text) Then
                    Console.WriteLine($"{vbCrLf}   Found Heading 1: '{text}'")

                    ' ---- Chapter with Heading 2 title(s) ----
                    If Regex.IsMatch(text, "^\s*chapter\b", RegexOptions.IgnoreCase) Then
                        Console.WriteLine($"   Processing chapter: '{text}'")

                        Dim chapter As String = text
                        Dim chapterTitleParts As New List(Of String)
                        Dim j As Integer = i + 1

                        Console.WriteLine($"   Looking for consecutive Heading 2 titles starting from paragraph {j}...")

                        Do While j <= total
                            Try
                                Dim p As Paragraph = doc.Paragraphs(j)
                                Dim pStyle As String = p.Style.NameLocal.ToLower()
                                Dim pOutlineLevel As WdOutlineLevel = p.OutlineLevel
                                Dim pIsHeading1 As Boolean = (pOutlineLevel = WdOutlineLevel.wdOutlineLevel1) OrElse pStyle.Contains("heading 1")
                                Dim pIsHeading2 As Boolean = (pOutlineLevel = WdOutlineLevel.wdOutlineLevel2) OrElse pStyle.Contains("heading 2")
                                Dim rawText As String = p.Range.Text
                                Dim cleanedText As String = CleanTitleText(rawText)

                                Console.WriteLine($"   Paragraph {j}: '{cleanedText}' (style: {pStyle})")

                                If String.IsNullOrWhiteSpace(rawText) Then
                                    Console.WriteLine($"   Empty paragraph {j}, skipping")
                                    j += 1
                                    Continue Do
                                End If

                                If pIsHeading2 Then
                                    chapterTitleParts.Add(cleanedText)
                                    Console.WriteLine($"   Collected Heading 2 part: '{cleanedText}'")
                                    j += 1
                                    Continue Do
                                ElseIf pIsHeading1 Then
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

                ElseIf isHeading2 AndAlso Not String.IsNullOrEmpty(text) Then
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

                Dim titleText As String = CStr(h("text"))
                ' Final safety: remove line breaks
                titleText = titleText.Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ")
                titleText = Regex.Replace(titleText, "\s+", " ").Trim()

                Console.WriteLine($"   TOC entry: '{titleText}' -> {h("page")}")

                ' Set text FIRST, then apply formatting (prevents paragraph style reset)
                para.Range.Text = $"{titleText}{vbTab}{h("page")}"

                ' Apply paragraph formatting AFTER setting text
                para.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
                para.Format.LeftIndent = 36
                para.Format.FirstLineIndent = -36
                ' Clear any inherited tab stops before adding ours
                para.Format.TabStops.ClearAll()
                para.Format.TabStops.Add(432,
                    WdTabAlignment.wdAlignTabRight,
                    WdTabLeader.wdTabLeaderDots)

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

                ' Remove existing page numbers
                Console.WriteLine($"   Removing {pageNums.Count} existing page numbers...")
                Do While pageNums.Count > 0
                    pageNums.Item(1).Delete()
                Loop

                ' Prefer robust field-code method for Roman numerals across Word interop versions.
                ' Keep restart settings where supported, then insert explicit Roman PAGE field.
                Try
                    pageNums.RestartNumberingAtSection = True
                    pageNums.StartingNumber = 1
                Catch
                    ' Some interop versions can be inconsistent here; field method below still applies Roman style.
                End Try

                footer.Range.Delete()
                Dim footerRange As WordRange = footer.Range
                footerRange.ParagraphFormat.Alignment = CType(wdAlignParagraphCenter, WdParagraphAlignment)

                Dim field As Field = footerRange.Fields.Add(
                    Range:=footerRange,
                    Type:=WdFieldType.wdFieldPage,
                    PreserveFormatting:=False
                )
                field.Code.Text = "PAGE \\* ROMAN \\* LOWER"
                field.Update()

                Console.WriteLine("   Roman numbering configured via PAGE field (i, ii, iii...)")

            Catch ex As Exception
                Console.WriteLine($"   ERROR: Page number setup failed: {ex.Message}")
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
                If sec2Footer.Range.Fields.Count > 0 Then
                    Dim codeText As String = sec2Footer.Range.Fields.Item(1).Code.Text.ToUpperInvariant()
                    If codeText.Contains("ROMAN") Then
                        Console.WriteLine("   Roman PAGE field verified!")
                    Else
                        Console.WriteLine("   Warning: Footer field is present but not Roman")
                    End If
                Else
                    Console.WriteLine("   Warning: No footer field found for verification")
                End If
            Catch ex As Exception
                Console.WriteLine($"   Warning: Verification failed: {ex.Message}")
            End Try

        Catch ex As Exception
            Console.WriteLine($"   ERROR: Pagination error: {ex.Message}")
        End Try
    End Sub

    ' -----------------------------------
    ' APPLY BOOK LAYOUT TO OUTPUT
    ' -----------------------------------
    Sub ApplyBookLayoutToOutput(outputDoc As Document,
                                bookDocPath As String,
                                Optional pageWidthTwipsOverride As Integer? = Nothing,
                                Optional pageHeightTwipsOverride As Integer? = Nothing)
        Console.WriteLine(vbCrLf & "Applying book layout (margins, page size, orientation...)")

        Try
            Dim word As Application = GetWord()
            Dim bookDoc As Document = word.Documents.Open(Path.GetFullPath(bookDocPath))

            Dim bookSection As Section = bookDoc.Sections.Item(1)
            Dim bookSetup As PageSetup = bookSection.PageSetup

            Dim leftMargin As Single = bookSetup.LeftMargin
            Dim rightMargin As Single = bookSetup.RightMargin
            Dim topMargin As Single = bookSetup.TopMargin
            Dim bottomMargin As Single = bookSetup.BottomMargin
            Dim pageWidth As Single = bookSetup.PageWidth
            Dim pageHeight As Single = bookSetup.PageHeight
            Dim orientation As WdOrientation = bookSetup.Orientation
            Dim gutter As Single = bookSetup.Gutter
            Dim mirrorMargins As Integer = bookSetup.MirrorMargins

            If pageWidthTwipsOverride.HasValue AndAlso pageHeightTwipsOverride.HasValue Then
                pageWidth = CSng(pageWidthTwipsOverride.Value / 20.0)
                pageHeight = CSng(pageHeightTwipsOverride.Value / 20.0)
                Console.WriteLine($"   Page size override from twips: {pageWidthTwipsOverride.Value} x {pageHeightTwipsOverride.Value}")
            End If

            Console.WriteLine("   Book layout detected:")
            Console.WriteLine($"      - Margins: L={leftMargin / 72.0:F2}in, R={rightMargin / 72.0:F2}in")
            Console.WriteLine($"                T={topMargin / 72.0:F2}in, B={bottomMargin / 72.0:F2}in")
            Console.WriteLine($"      - Page size: {pageWidth / 72.0:F2}"" x {pageHeight / 72.0:F2}"")
            Console.WriteLine($"      - Orientation: {If(orientation = WdOrientation.wdOrientLandscape, "Landscape", "Portrait")}")

            bookDoc.Close(False)
            word.Quit()

            Dim outputSections As Sections = outputDoc.Sections
            Console.WriteLine($"   Applying to {outputSections.Count} sections...")

            For i As Integer = 1 To outputSections.Count
                Dim sec As Section = outputSections.Item(i)
                Dim setup As PageSetup = sec.PageSetup

                setup.LeftMargin = leftMargin
                setup.RightMargin = rightMargin
                setup.TopMargin = topMargin
                setup.BottomMargin = bottomMargin
                setup.Orientation = orientation
                ' Set width/height AFTER orientation so Word does not auto-resize to a smaller preset.
                setup.PageWidth = pageWidth
                setup.PageHeight = pageHeight
                setup.Gutter = gutter
                setup.MirrorMargins = mirrorMargins

                Dim effectiveWidthTwips As Integer = CInt(Math.Round(setup.PageWidth * 20.0))
                Dim effectiveHeightTwips As Integer = CInt(Math.Round(setup.PageHeight * 20.0))
                Console.WriteLine($"      Section {i}: layout applied (headers/footers preserved)")
                Console.WriteLine($"         Effective size: {effectiveWidthTwips} x {effectiveHeightTwips} twips")
            Next

            Console.WriteLine("   Book layout successfully applied to all sections")

        Catch ex As Exception
            Console.WriteLine($"   WARNING: Layout application error: {ex.Message}")
        End Try
    End Sub

    ' -----------------------------------
    ' ASSEMBLE FINAL DOCUMENT
    ' -----------------------------------
    Sub AssembleFinal(titleDoc As String,
                      copyrightDoc As String,
                      tocDoc As String,
                      output As String,
                      Optional bookDocPath As String = Nothing,
                      Optional pageWidthTwipsOverride As Integer? = Nothing,
                      Optional pageHeightTwipsOverride As Integer? = Nothing,
                      Optional applyBookLayout As Boolean = True)
        Console.WriteLine(vbCrLf & "Assembling final document with Roman pagination fix")

        Dim word As Application = GetWord()
        Dim doc As Document = word.Documents.Add()

        Dim sel As Selection = word.Selection

        Try
            ' Title page (Section 1) - inserted via Selection.InsertFile
            ' (equivalent to Word UI: Insert > Object > Text from File)
            If File.Exists(titleDoc) Then
                Console.WriteLine("   Inserting title page (Section 1)...")
                sel.EndKey(WdUnits.wdStory) ' cursor to end of doc
                sel.InsertFile(FileName:=Path.GetFullPath(titleDoc))

                sel.EndKey(WdUnits.wdStory)
                sel.InsertBreak(WdBreakType.wdSectionBreakNextPage)
                Console.WriteLine("   Title page + section break inserted")
            End If

            ' Copyright page (Section 2, page "i")
            If File.Exists(copyrightDoc) Then
                Console.WriteLine("   Inserting copyright page (will be Roman 'i')...")
                sel.EndKey(WdUnits.wdStory)
                sel.InsertFile(FileName:=Path.GetFullPath(copyrightDoc))

                sel.EndKey(WdUnits.wdStory)
                sel.InsertBreak(WdBreakType.wdPageBreak)
                Console.WriteLine("   Copyright page inserted")
            End If

            ' TOC (Section 2, pages "ii", "iii", ...)
            Console.WriteLine("   Inserting TOC (will be Roman 'ii', 'iii'...)...")
            sel.EndKey(WdUnits.wdStory)
            sel.InsertFile(FileName:=Path.GetFullPath(tocDoc))
            Console.WriteLine("   TOC inserted")

            ' Apply Roman pagination
            ApplyRomanPagination(doc)

            If applyBookLayout AndAlso Not String.IsNullOrWhiteSpace(bookDocPath) Then
                ApplyBookLayoutToOutput(doc,
                                        bookDocPath,
                                        pageWidthTwipsOverride,
                                        pageHeightTwipsOverride)
            End If

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
        Dim cfg As BuilderConfig = LoadBuilderConfig(projectRoot)

        Directory.CreateDirectory(cfg.OutputDir)
        Directory.CreateDirectory(cfg.TempDir)

        Dim titleFile As String = cfg.TitleFile
        Dim copyrightFile As String = cfg.CopyrightFile
        Dim bookFile As String = cfg.BookBodyFile
        Dim outputFile As String = Path.Combine(cfg.OutputDir, cfg.OutputFilename)
        Dim tocTemp As String = Path.Combine(cfg.TempDir, "toc.docx")

        Console.WriteLine($"{vbCrLf}Project root: {projectRoot}")
        Console.WriteLine($"Config: {MakeRelative(projectRoot, cfg.ConfigPath)}")
        Console.WriteLine($"Book: {MakeRelative(projectRoot, bookFile)}")
        Console.WriteLine($"Title: {MakeRelative(projectRoot, titleFile)}")
        Console.WriteLine($"Copyright: {MakeRelative(projectRoot, copyrightFile)}")
        Console.WriteLine($"Temp TOC: {MakeRelative(projectRoot, tocTemp)}")
        Console.WriteLine($"Output: {MakeRelative(projectRoot, outputFile)}")
        If cfg.PageWidthTwipsOverride.HasValue AndAlso cfg.PageHeightTwipsOverride.HasValue Then
            Console.WriteLine($"Page size override (twips): {cfg.PageWidthTwipsOverride.Value} x {cfg.PageHeightTwipsOverride.Value}")
        End If

        Dim headings = ExtractHeadings(bookFile)

        If headings.Count > 0 Then
            BuildTocDoc(headings, tocTemp)
            AssembleFinal(titleFile,
                          copyrightFile,
                          tocTemp,
                          outputFile,
                          bookFile,
                          cfg.PageWidthTwipsOverride,
                          cfg.PageHeightTwipsOverride,
                          cfg.ApplyBookLayout)

            If cfg.DeleteTempToc AndAlso File.Exists(tocTemp) Then
                File.Delete(tocTemp)
            End If
        End If
    End Sub

End Module
