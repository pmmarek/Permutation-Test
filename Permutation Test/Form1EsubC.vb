Imports System
Imports System.IO
Imports System.Diagnostics
Imports System.ComponentModel
Imports System.Windows.Forms
Imports Microsoft.VisualBasic.FileIO
Imports Microsoft.Office.Interop

Public Class frmPermTest
   ' This program produces a formatted list of HSYAS endpoint analyses.
   ' It reads a file of endpoint variables, prints a summary of the
   ' variables, applies a permutation test to the variables, prints
   ' p-values from the test and produces a confidence interval for
   ' effect size.  It does this for each endpoint within each
   ' gender group and for combined male + female data.

   ' THIS VERSION WAS DESIGNED WITH THE HSHSS EVALUATION EFFORT IN MIND.
   ' IT USES A CLUSTER-SIZE WEIGHTED PERMUTATION TEST STATISTIC ON 
   ' 25 PAIRS OF GROUP-AGGREGATED DATA.  IT HANDLES INTEGER OR FLOATING
   ' POINT ENDPOINTS (I.E., IT COMBINES THE TWO HSPP VERSIONS).

   ' PROJECT FOLDER: C:\Users\pmarek\Documents\_
   '               Visual Studio 20105\Projects\HS Permutation Test\_
   '               Permutation Test\
   ' PROJECT FILE: <PROJECT FOLDER>\HS Permutation Test.vbproj
   ' RELEASE SOURCE: <PROJECT FOLDER>\Form1.vb
   ' RELEASE BINARY: <PROJECT FOLDER>\bin\Release\HSPermutationTest.exe
   ' PUBLISHING FOLDER: \\phs.fhcrc.org\Projects\HSPP\SHARED\SETUP\VS2010\HSPermTest

   ' Inputs: An ASCII input file of aggregated endpoint data, produced by
   '         an SPSS program.
   ' Outputs: 1 - A Microsoft Word document file with a detailed list of
   '              inputs, results and p-values, which should be stored
   '              for future reference.
   '          2 - A Microsoft Word document file with a one-line summary of
   '              results and p-values for each gender group in each run.

   ' MAJOR SUBROUTINES:
   '    mnuRun_Click - Handler for the Run command on the menu strip; it
   '                   performs many of the functions of the main program
   '                   in the original FORTRAN version of the program.
   '    ReadData     - Reads and parses the ASCII input data file.
   '    PermPval     - Performs a permutation test and returns p-values.
   '    PermCI       - Returns a 95% permutation confidence interval
   '                   for treatment effect.
   '    Detail_out   - Produces a Word document file of detailed results.
   '    Summary_out  - Produces a Word document file of summary results.

   ' AUTHOR: Pat Marek
   ' DATE:   12/21/2007 (based on an earlier FORTRAN version written by Pat Marek
   '    for the HSPP endpoint analyses of 2000, which was in turn based on 
   '    versions of 1992-94 written by Lynn Onstad and Pat Marek)
   ' REVISION HISTORY
   ' #    Date    Details
   ' -  --------  -------
   ' 1  01/09/08  Disabled Run command after it is clicked to avoid
   '              multiple executions on same dataset.  Added more
   '              error handling to Word manipulations.
   ' 2  03/06/08  Changed the difference (delta) computations to use
   '              (Experimental - Control), to accommodate a decision to use
   '              endpoints where positive values are "good" (unlike HSPP).
   ' 3  05/12/14  Changed Namespace declarations and Word VBA to work with Visual
   '              Studio 2008/2010 and Office 2010.  Changed default extension for
   '              saving results to DOCX, and changed FileOpen dialog to look for
   '              DOC or DOCX.  Corrected extra paragraph marks that are inserted
   '              after page breaks in Word 2010.

   Const HSPPGender As Boolean = False  ' HSPP had 1=Male, 2=Female; HSHSS is reversed

   Dim strDataFile As String, strResultsFile As String
   Dim resltc(2, 24) As Double, reslte(2, 24) As Double, resltct(2) As Double, _
      resltet(2) As Double
   Dim d(24) As Double, dobs0 As Double, StartTime As Double
   Dim dmc(24) As Double, dme(24) As Double, dtotmc As Double, dtotme As Double
   Dim nsqc(2, 24) As Integer, nsqe(2, 24) As Integer, nsqct(2) As Integer, _
      nsqet(2) As Integer
   Dim notapc(2, 24) As Integer, notape(2, 24) As Integer, notapct(2) As Integer, _
      notapet(2) As Integer
   Dim m As Integer, imc(24) As Integer, ime(24) As Integer, inc(24) As Integer, _
      ine(24) As Integer
   Dim validc(2, 24) As Integer, valide(2, 24) As Integer, smissc(2, 24) As Integer, _
      smisse(2, 24) As Integer
   Dim validct(2) As Integer, validet(2) As Integer, smissct(2) As Integer, _
      smisset(2) As Integer
   Dim excluc(2, 24) As Integer, exclue(2, 24) As Integer, excluct(2) As Integer, _
      excluet(2) As Integer
   Dim sex As Integer, varnum As Integer, mainendp As Integer, floatpt As Integer
   Dim totnc As Integer, totne As Integer, totmc As Integer, totme As Integer
   Dim varname As String, vardescr As String, category As String, group As String, _
      subgroup As String, bolWordAlreadyOpen As Boolean, bolFloatPt As Boolean
   Dim appWord As Word.Application

   Private Sub mnuFileOpenData_Click(ByVal sender As Object, _
      ByVal e As System.EventArgs) Handles mnuFileOpenData.Click
      ' This routine handles the menu's File/Open/Data choice.  It must
      ' be called prior to running an analysis.
      Dim dr As DialogResult
      dlgOpen.Filter = "Data files (*.dat)|*.dat"
      dlgOpen.InitialDirectory = "N:\HSPP\Data\HSHSS\HSYAS\Eval\PermTests\"
      dr = dlgOpen.ShowDialog()
      If dr = Windows.Forms.DialogResult.OK Then
         strDataFile = dlgOpen.FileName
         If LCase(Microsoft.VisualBasic.Right(strDataFile, 4)) <> ".dat" Then
            MessageBox.Show("Data files must have the extension '.dat'.", _
               "Not a data file", MessageBoxButtons.OK, _
            MessageBoxIcon.Exclamation)
            strDataFile = ""
         End If
      Else
         MessageBox.Show("Note: You must open a data file before running the " & _
            "Permutation Test.", "Open Data", MessageBoxButtons.OK, _
            MessageBoxIcon.Information)
      End If
   End Sub

   Private Sub mnuFileOpenResults_Click(ByVal sender As Object, _
      ByVal e As System.EventArgs) Handles mnuFileOpenResults.Click
      ' This routine handles the menu's File/Open/Results choice.  It gives
      ' the user a way to open an existing results document.
      Dim dr As DialogResult
      dlgOpen.Filter = "Result documents (*.doc, *.docx)|*.doc;*.docx"
      dlgOpen.InitialDirectory = "N:\HSPP\Data\HSHSS\HSYAS\Eval\PermTests\"
      dr = dlgOpen.ShowDialog()
      If dr = Windows.Forms.DialogResult.OK Then
         strResultsFile = dlgOpen.FileName
         If LCase(Microsoft.VisualBasic.Right(strResultsFile, 4)) <> ".doc" _
            And LCase(Microsoft.VisualBasic.Right(strResultsFile, 5)) <> ".docx" Then
            MessageBox.Show("The results file must have the Word extension '.doc' or '.docx'.", _
               "Not a Word file", MessageBoxButtons.OK, _
            MessageBoxIcon.Exclamation)
            strResultsFile = ""
         End If
      Else
         MessageBox.Show("Note: You must open a results file before printing " & _
            "results.", "Open Results", MessageBoxButtons.OK, _
            MessageBoxIcon.Information)
      End If
   End Sub

   Private Sub mnuFilePrint_Click(ByVal sender As Object, _
      ByVal e As System.EventArgs) Handles mnuFilePrint.Click
      ' This routine handles the menu's File/Print choice.  If the user has
      ' previously opened a results file, this code will print it.
      If strResultsFile = "" Then
         MessageBox.Show("You must open a results file before printing " & _
                     "results.", "Print Results", MessageBoxButtons.OK, _
                     MessageBoxIcon.Exclamation)
      Else
         ' Use Process.Start to automatically print the results file.  This
         ' produces the same result as calling the ShellExecute function from
         ' the Windows API
         Dim MyProcess As New Process
         MyProcess.StartInfo.FileName = strResultsFile
         MyProcess.StartInfo.Verb = "Print"
         MyProcess.StartInfo.CreateNoWindow = True
         MyProcess.Start()
      End If
   End Sub

   Private Sub mnuRun_Click(ByVal sender As Object, _
      ByVal e As System.EventArgs) Handles mnuRun.Click
      ' This is the main routine for running the I/O and analysis procedures 
      ' of the Permutation Test; it is roughly equivalent to the main program
      ' of the old FORTRAN code.
      Dim pageno As Integer, spageno As Integer, slineno As Integer
      Dim usepair(24) As Boolean
      Dim p1 As Double, p2 As Double, dl As Double, du As Double
      Dim r1pc As Double
      ' Data reader formats; these represent field widths
      Dim Format1() As Integer = {12, 40, 2, 3, 8, 1, 1, 60, 84, -1}
      Dim Format2(26) As Integer
      Dim Format3(26) As Integer
      Dim intI As Integer, intSex As Integer, dr As DialogResult
      Dim strSexLabel() As String = {"Females", "Males", "Combined"}

      ' Check that a data file was specified.
      If strDataFile = "" Then
         MessageBox.Show("You must open a data file before running " & _
            "the Permutation Test.", "Run Test", MessageBoxButtons.OK, _
            MessageBoxIcon.Exclamation)
         Exit Sub
      End If
      ' Good to go: Disable run command, to avoid multiple executions
      ' on the current dataset.
      mnuRun.Enabled = False
      ' Set start time, to be used in elapsed time display
      StartTime = DateAndTime.Timer
      ' Declare Word object variables and names of the summary and detailed
      ' output documents.  Assume that they will have the same location
      ' as the data file, but with "_detail" and "_summary" added to the file
      ' name and standard Word extension ".doc".
      Dim strSummarydoc As String = _
         Replace(strDataFile, ".dat", "_summary.doc", , , CompareMethod.Text)
      Dim strDetaildoc As String = _
         Replace(strDataFile, ".dat", "_detail.doc", , , CompareMethod.Text)
      Dim docSummary As Word.Document
      Dim docDetail As Word.Document
      Dim objRange As Word.Range

      ' Initialize data read formats.
      For intI = 0 To 25
         Format2(intI) = 10
         Format3(intI) = 5
      Next
      Format2(26) = -1
      Format3(26) = -1
      ' Reset the progress window
      txtProgress.Clear()

      ' Open Word (or a new copy of Word) and add two new documents.  Note that
      ' only one instance of Word will be created, even though two documents
      ' are opened.
      Try
         appWord = GetObject(, "Word.Application")
         bolWordAlreadyOpen = True
      Catch exc As Exception
         appWord = CreateObject("Word.Application")
         bolWordAlreadyOpen = False
      End Try

      docDetail = appWord.Documents.Add
      docSummary = appWord.Documents.Add

      ' Set to Word 2007 compatibility mode -- otherwise anchoring of graphic lines
      ' won't work due to a Word 2010 VBA bug.
      docDetail.SetCompatibilityMode(Word.WdCompatibilityMode.wdWord2007)
      docSummary.SetCompatibilityMode(Word.WdCompatibilityMode.wdWord2007)

      ' Page setup parameters for detail output (convert inches to points)
      With docDetail.PageSetup
         .LeftMargin = 0.66 * 72
         .RightMargin = 0.5 * 72
         .TopMargin = 0.75 * 72
         .BottomMargin = 0.75 * 72
      End With
      ' Set font to Arial, 11 pt.
      objRange = docDetail.Content
      With objRange.Font
         .Name = "Arial"
         .Size = 11
      End With
      ' Set default paragraph spacing
      With docDetail.Paragraphs
         .LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly
         .LineSpacing = 11
         .SpaceBefore = 0
         .SpaceAfter = 0
      End With

      ' Page setup parameters for summary output (in points)
      With docSummary.PageSetup
         .Orientation = Word.WdOrientation.wdOrientLandscape
         .LeftMargin = 72
         .RightMargin = 72
         .TopMargin = 36
         .BottomMargin = 36
      End With
      ' Set font to Arial, 11 pt.
      objRange = docSummary.Content
      With objRange.Font
         .Name = "Arial"
         .Size = 11
      End With
      ' Set default paragraph spacing
      With docSummary.Paragraphs
         .LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle
         .SpaceBefore = 0
         .SpaceAfter = 0
      End With

      ' Set up the reader object in a Using block, which contains the main
      ' read/analyze/report loop.
      Using DataReader As New TextFieldParser(strDataFile)
         DataReader.TextFieldType = FieldType.FixedWidth
         Me.UseWaitCursor = True
         While Not DataReader.EndOfData
            ' The data for each variable is paired, with the data for females
            ' (intSex=1) first and the data for males (intSex=2) second. The
            ' combined data (intSex=3) for each variable is calculated from 
            ' the gender-specific data; ReadData reads or computes the data.
            For intSex = 1 To 3
               ReadData(Format1, Format2, Format3, DataReader, intSex)
               ' Determine which school pairs should be involved in the test.
               ' In order to qualify, each member of the pair must have at
               ' least one record with useable data for the endpoint (i.e.,
               ' neither missing nor excluded).

               ' This condition is satisfied for the ith pair if validc(i) > 0
               ' and valide(i) > 0.

               m = 0
               For intI = 0 To 24
                  If (validc(intSex - 1, intI) > 0 And valide(intSex - 1, intI) > 0) Then
                     usepair(intI) = True
                     inc(m) = validc(intSex - 1, intI)
                     ine(m) = valide(intSex - 1, intI)
                     If bolFloatPt Then
                        dmc(m) = resltc(intSex - 1, intI) * validc(intSex - 1, intI)
                        dme(m) = reslte(intSex - 1, intI) * valide(intSex - 1, intI)
                     Else
                        ' CInt rounds its argument to the nearest integer and returns an integer.
                        ' This ensures the correct integer result for the number of individuals
                        ' with outcome value 1 in the control and experimental members of each pair.
                        imc(m) = CInt(resltc(intSex - 1, intI) * validc(intSex - 1, intI))
                        ime(m) = CInt(reslte(intSex - 1, intI) * valide(intSex - 1, intI))
                     End If
                     m = m + 1
                  Else
                     usepair(intI) = False
                  End If
               Next
               ' If there is at least one school pair satisfying the
               ' conditions, perform the permutation test.  Set the total
               ' counts into totmc, totnc, totnc and totne.  d(i) is the
               ' unweighted prevalence difference in each cluster, which will
               ' be used to produce starting values for the confidence
               ' interval procedure.  dobs0 is the test statistic for a test
               ' of H0: delta = 0.

               If m > 0 Then
                  totmc = 0
                  totme = 0
                  dtotmc = 0.0
                  dtotme = 0.0
                  totnc = 0
                  totne = 0
                  For intI = 0 To m - 1
                     If bolFloatPt Then
                        dtotmc = dtotmc + dmc(intI)
                        dtotme = dtotme + dme(intI)
                        d(intI) = dme(intI) / CDbl(ine(intI)) _
                           - dmc(intI) / CDbl(inc(intI))
                     Else
                        totmc = totmc + imc(intI)
                        totme = totme + ime(intI)
                        d(intI) = CDbl(ime(intI)) / CDbl(ine(intI)) _
                           - CDbl(imc(intI)) / CDbl(inc(intI))
                     End If
                     totnc = totnc + inc(intI)
                     totne = totne + ine(intI)
                  Next
                  If bolFloatPt Then
                     dobs0 = dtotme / CDbl(totne) - dtotmc / CDbl(totnc)
                  Else
                     dobs0 = CDbl(totme) / CDbl(totne) _
                        - CDbl(totmc) / CDbl(totnc)
                  End If
                  PermPval(p1, p2, 0.0)
                  PermCI(dl, du)
               End If

               ' Compute percent increase. -1.0e+6 is used as a flag for
               ' cases where the % increase doesn't exist.
               If m > 0 And resltct(intSex - 1) <> 0 And validct(intSex - 1) > 0 Then
                  r1pc = (resltet(intSex - 1) - resltct(intSex - 1)) _
                     / resltct(intSex - 1) * 100
               Else
                  r1pc = -1000000.0
               End If
               ' Output a one-page detailed listing for this group and endpoint
               detail_out(docDetail, strSexLabel(intSex - 1), pageno, m, usepair, r1pc, dobs0, p2, dl, du)
               ' Output a one-line summary for this group and endpoint
               summary_out(docSummary, strSexLabel(intSex - 1), spageno, slineno, m, r1pc, dobs0, p2, dl, du)
               ' Write a one-line progress notice to the main form.
               txtProgress.AppendText(CStr(IIf(pageno > 1, vbCrLf, "")) & _
                  "Finished Endpoint " & CStr(varnum) & " (" & varname & ") " & _
                  strSexLabel(intSex - 1))
               txtProgress.ScrollToCaret()
               ' Force screen updates to appear *now*
               Application.DoEvents()
            Next
         End While
         ' Show elapsed timeon main form
         txtProgress.AppendText(vbCrLf & "Elapsed time: " & _
            CStr(DateAndTime.Timer - StartTime) & " seconds")
         Me.UseWaitCursor = False
      End Using
      ' Save the detail and summary results documents.  In each case, check to
      ' see if there is already a results file with the same name; if so, give
      ' the user the chance to save it under whatever path and name they desire.
      If File.Exists(strDetaildoc) Then
         dr = MessageBox.Show("The file " & strDetaildoc & _
            " already exists." & vbCrLf & "Do you want to overwrite it with the results" & _
            " of this run?", "Detailed Output File Exists", MessageBoxButtons.YesNo, _
            MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
         If dr = Windows.Forms.DialogResult.No Then
            MessageBox.Show("You will be prompted for a location and " & _
               "name for saving the detailed results.", "", MessageBoxButtons.OK, _
               MessageBoxIcon.Information)
            Try
               docDetail.Save()
            Catch exc As Exception
               MessageBox.Show("Failed to save detailed results document " & _
                  strDetaildoc & vbCrLf & exc.Message, "Problem Saving File", _
                  MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End Try
         Else
            Try
               docDetail.SaveAs(FileName:=strDetaildoc)
               MessageBox.Show("Detailed results were saved automatically to " & _
                  strDetaildoc & "," & vbCrLf & "overwriting the previous version.", _
                  "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch exc As Exception
               MessageBox.Show("Failed to save detailed results document " & _
                  strDetaildoc & vbCrLf & exc.Message, "Problem Saving File", _
                  MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End Try
         End If
      Else  'Detailed results file doesn't exist; save document
         Try
            docDetail.SaveAs(FileName:=strDetaildoc)
            MessageBox.Show("Detailed results were saved automatically to " & _
               strDetaildoc, "", MessageBoxButtons.OK, MessageBoxIcon.Information)
         Catch exc As Exception
            MessageBox.Show("Failed to save detailed results document " & _
               strDetaildoc & vbCrLf & exc.Message, "Problem Saving File", _
               MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
         End Try
      End If
      If File.Exists(strSummarydoc) Then
         dr = MessageBox.Show("The file " & strSummarydoc & _
            " already exists." & vbCrLf & "Do you want to overwrite it with the results" & _
            " of this run?", "Summary Output File Exists", MessageBoxButtons.YesNo, _
            MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
         If dr = Windows.Forms.DialogResult.No Then
            MessageBox.Show("You will be prompted for a location and " & _
               "name for saving the summary results.", "", MessageBoxButtons.OK, _
               MessageBoxIcon.Information)
            Try
               docSummary.Save()
            Catch exc As Exception
               MessageBox.Show("Failed to save summary results document " & _
                  strSummarydoc & vbCrLf & exc.Message, "Problem Saving File", _
                  MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End Try
         Else
            Try
               docSummary.SaveAs(FileName:=strSummarydoc)
               MessageBox.Show("Summary results were saved automatically to " & _
                  strSummarydoc & "," & vbCrLf & "overwriting the previous version.", _
                  "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch exc As Exception
               MessageBox.Show("Failed to save summary results document " & _
                  strSummarydoc & vbCrLf & exc.Message, "Problem Saving File", _
                  MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End Try
         End If
      Else  'Summary results file doesn't exist; save document
         Try
            docSummary.SaveAs(FileName:=strSummarydoc)
            MessageBox.Show("Summary results were saved automatically to " & _
               strSummarydoc, "", MessageBoxButtons.OK, MessageBoxIcon.Information)
         Catch exc As Exception
            MessageBox.Show("Failed to save summary results document " & _
               strSummarydoc & vbCrLf & exc.Message, "Problem Saving File", _
               MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
         End Try
      End If
      ' Give the user the option of printing either results document now.
      dr = MessageBox.Show("Would you like to print the detailed results to " & _
         "your default printer now?", "Print Detailed Results", MessageBoxButtons.YesNo, _
         MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
      If dr = Windows.Forms.DialogResult.Yes Then
         Try
            docDetail.PrintOut()
            MessageBox.Show("Detailed results were sent to your default printer.", _
               "", MessageBoxButtons.OK, MessageBoxIcon.Information)
         Catch exc As Exception
            MessageBox.Show("Failed to print detailed results document " & _
               strDetaildoc & vbCrLf & exc.Message, "Printing Problem", _
               MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
         End Try
      End If
      dr = MessageBox.Show("Would you like to print the summary results to " & _
         "your default printer now?", "Print Summary Results", MessageBoxButtons.YesNo, _
         MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
      If dr = Windows.Forms.DialogResult.Yes Then
         Try
            docSummary.PrintOut()
            MessageBox.Show("Summary results were sent to your default printer.", _
               "", MessageBoxButtons.OK, MessageBoxIcon.Information)
         Catch exc As Exception
            MessageBox.Show("Failed to print summary results document " & _
               strSummarydoc & vbCrLf & exc.Message, "Printing Problem", _
               MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
         End Try
      End If
      ' Add a message to alert the user that there is no more processing for this run.
      MessageBox.Show("All processing for this run has been completed.", _
         "Finished", MessageBoxButtons.OK, MessageBoxIcon.Information)
      ' Reset strDataFile in case the user wants to do more runs.
      strDataFile = ""
      ' Close document windows and exit Word, unless the user had it open prior
      ' to starting the permutation test program.  The "do not save" option is
      ' used because the documents were saved previously.  The action of printing
      ' a document sometimes makes a change to header information in Word; if we
      ' didn't use this option, the user might get extra prompts to save.
      If Not docDetail Is Nothing Then
         Try
            docDetail.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
         Catch exc As Exception
            MessageBox.Show("Failed to close detailed results document " & _
               strSummarydoc & vbCrLf & exc.Message, "Document not closed", _
               MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
         End Try
      End If
      If Not docSummary Is Nothing Then
         Try
            docSummary.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
         Catch exc As Exception
            MessageBox.Show("Failed to close summary results document " & _
               strSummarydoc & vbCrLf & exc.Message, "Document not closed", _
               MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
         End Try
      End If
      ' If a new copy of word was created, close it
      If Not bolWordAlreadyOpen And Not appWord Is Nothing Then
         Try
            appWord.Quit()
            appWord = Nothing
         Catch exc As Exception
            MessageBox.Show("Failed to exit Word;  " & _
               "you may need to close it manually." & vbCrLf & _
               exc.Message, "MS Word not closed", _
               MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
         End Try
      End If
      ' Re-enable the run command
      mnuRun.Enabled = True
   End Sub

   Private Sub ReadData(ByRef Format1() As Integer, ByRef Format2() As Integer, _
      ByRef Format3() As Integer, ByRef MyReader As TextFieldParser, ByVal intSex As Integer)
      ' This procedure reads the data for a single variable and gender.  The
      ' calling routine (mnuRun_Click) opened the data file as a TextFieldParser
      ' object, which has methods for dealing with fixed-format text data
      ' records.  The fields are read into a string array, then parsed to
      ' program variables.

      ' Explanation of variables read from the data file:
      ' ---------------------------------------------------------------------
      ' GROUP:    Group of interest for this analysis
      ' SUBGROUP: Subgroup of interest for this analysis
      ' SEX:      Gender (1=Female, 2=Male)
      ' VARNUM:   Variable number
      ' VARNAME:  Variable name (max length 8 characters)
      ' MAINENDP: Main endpoint? (1=Yes, 0=No)
      ' FLOATPT:  Floating point endpoint? (1=Yes, 0=No)
      ' CATEGORY: Category of variable
      ' VARDESCR: Variable description (up to 84 characters)

      ' RESLTC(SEX-1,I-1): Endpoint prevalence for pair I control, gender=SEX
      ' RESLTCT(SEX-1):    Endpoint prevalence over all controls, gender=SEX
      ' RESLTE(SEX-1,I-1): Endpoint prevalence for pair I exper., gender=SEX
      ' RESLTET(SEX-1):    Endpoint prevalence over all exper., gender=SEX
      ' NSQC(SEX-1,I-1):   Number of SQs for pair I controls, gender=SEX
      ' NSQCT(SEX-1):      Number of SQs over all controls, gender=SEX
      ' NSQE(SEX-1,I-1):   Number of SQs for pair I experimentals, gender=SEX
      ' NSQET(SEX-1):      Number of SQs over all experimentals, gender=SEX
      ' EXCLUC(SEX-1,I-1): Number of SQs excluded from analysis due to suspect
      '                      data in pair I controls, gender=SEX
      ' EXCLUCT(SEX-1):    Number of SQs excluded from analysis due to suspect
      '                      data over all controls, gender=SEX
      ' EXCLUE(SEX-1,I-1): Number of SQs excluded from analysis due to suspect
      '                      data in pair I experimentals, gender=SEX
      ' EXCLUET(SEX-1):    Number of SQs excluded from analysis due to suspect
      '                      data in all experimentals, gender=SEX
      ' NOTAPC(SEX-1,I-1): Number of SQs not applicable to analysis of this
      '                      endpoint in pair I controls, gender=SEX
      ' NOTAPCT(SEX-1):    Number of SQs not applicable to analysis of this
      '                      endpoint over all controls, gender=SEX
      ' NOTAPE(SEX-1,I-1): Number of SQs not applicable to analysis of this
      '                      endpoint in pair I experimentals, gender=SEX
      ' NOTAPET(SEX-1):    Number of SQs not applicable to analysis of this
      '                      endpoint in all experimentals, gender=SEX
      ' SMISSC(SEX-1,I-1): Number of SQs that lack the item(s) needed to define
      '                      this endpoint in pair I controls, gender=SEX
      ' SMISSCT(SEX-1):    Number of SQs that lack the item(s) needed to define
      '                      this endpoint over all controls, gender=SEX
      ' SMISSE(SEX-1,I-1): Number of SQs that lack the item(s) needed to define
      '                      this endpoint in pair I experimentals, gender=SEX
      ' SMISSET(SEX-1):    Number of SQs that lack the item(s) needed to define
      '                      this endpoint in all experimentals, gender=SEX
      ' VALIDC(SEX-1,I-1): Number of SQs that contributed data for this
      '                      endpoint in pair I controls, gender=SEX
      ' VALIDCT(SEX-1):    Number of SQs that contributed data for this
      '                      endpoint over all controls, gender=SEX
      ' VALIDE(SEX-1,I-1): Number of SQs that contributed data for this
      '                      endpoint in pair I experimentals, gender=SEX
      ' VALIDET(SEX-1):    Number of SQs that contributed data for this
      '                      endpoint in all experimentals, gender=SEX

      Dim CurrentLine() As String, intI As Integer
      If intSex < 3 Then ' Read data for females or males.
         ' The data file contains separate data in separate sets of rows for
         ' females (intSex=1) and males (intSex=2), for each variable.  The 
         ' Data for males+females combined is calculated rather than read.
         If Not MyReader.EndOfData Then
            ' Line 1
            MyReader.SetFieldWidths(Format1)
            CurrentLine = MyReader.ReadFields
            ' Parse out the fields into program variables
            group = CurrentLine(0)
            subgroup = CurrentLine(1)
            sex = Integer.Parse(CurrentLine(2))
            ' In HSPP data, 1=Male and 2=Female; if this coding is used, switch
            If HSPPGender Then sex = 3 - sex
            varnum = Integer.Parse(CurrentLine(3))
            varname = CurrentLine(4)
            mainendp = Integer.Parse(CurrentLine(5))
            floatpt = Integer.Parse(CurrentLine(6))
            bolFloatPt = (floatpt = 1)
            category = CurrentLine(7)
            vardescr = CurrentLine(8)
            ' Lines 2 - 3
            MyReader.SetFieldWidths(Format2)
            CurrentLine = MyReader.ReadFields
            For intI = 0 To 24
               resltc(sex - 1, intI) = Double.Parse(CurrentLine(intI))
            Next
            resltct(sex - 1) = Double.Parse(CurrentLine(25))
            CurrentLine = MyReader.ReadFields
            For intI = 0 To 24
               reslte(sex - 1, intI) = Double.Parse(CurrentLine(intI))
            Next
            resltet(sex - 1) = Double.Parse(CurrentLine(25))
            ' Lines 4 - 13
            MyReader.SetFieldWidths(Format3)
            CurrentLine = MyReader.ReadFields
            For intI = 0 To 24
               nsqc(sex - 1, intI) = Integer.Parse(CurrentLine(intI))
            Next
            nsqct(sex - 1) = Integer.Parse(CurrentLine(25))
            CurrentLine = MyReader.ReadFields
            For intI = 0 To 24
               nsqe(sex - 1, intI) = Integer.Parse(CurrentLine(intI))
            Next
            nsqet(sex - 1) = Integer.Parse(CurrentLine(25))
            CurrentLine = MyReader.ReadFields
            For intI = 0 To 24
               excluc(sex - 1, intI) = Integer.Parse(CurrentLine(intI))
            Next
            excluct(sex - 1) = Integer.Parse(CurrentLine(25))
            CurrentLine = MyReader.ReadFields
            For intI = 0 To 24
               exclue(sex - 1, intI) = Integer.Parse(CurrentLine(intI))
            Next
            excluet(sex - 1) = Integer.Parse(CurrentLine(25))
            CurrentLine = MyReader.ReadFields
            For intI = 0 To 24
               notapc(sex - 1, intI) = Integer.Parse(CurrentLine(intI))
            Next
            notapct(sex - 1) = Integer.Parse(CurrentLine(25))
            CurrentLine = MyReader.ReadFields
            For intI = 0 To 24
               notape(sex - 1, intI) = Integer.Parse(CurrentLine(intI))
            Next
            notapet(sex - 1) = Integer.Parse(CurrentLine(25))
            CurrentLine = MyReader.ReadFields
            For intI = 0 To 24
               smissc(sex - 1, intI) = Integer.Parse(CurrentLine(intI))
            Next
            smissct(sex - 1) = Integer.Parse(CurrentLine(25))
            CurrentLine = MyReader.ReadFields
            For intI = 0 To 24
               smisse(sex - 1, intI) = Integer.Parse(CurrentLine(intI))
            Next
            smisset(sex - 1) = Integer.Parse(CurrentLine(25))
            CurrentLine = MyReader.ReadFields
            For intI = 0 To 24
               validc(sex - 1, intI) = Integer.Parse(CurrentLine(intI))
            Next
            validct(sex - 1) = Integer.Parse(CurrentLine(25))
            CurrentLine = MyReader.ReadFields
            For intI = 0 To 24
               valide(sex - 1, intI) = Integer.Parse(CurrentLine(intI))
            Next
            validet(sex - 1) = Integer.Parse(CurrentLine(25))
         End If
      Else
         ' Calculate combined male+female data when intSex=3;
         ' note that the variables that are not calculated here
         ' will be unchanged from the previously read values.
         sex = 3
         For intI = 0 To 24
            nsqc(2, intI) = nsqc(0, intI) + nsqc(1, intI)
            nsqe(2, intI) = nsqe(0, intI) + nsqe(1, intI)
            excluc(2, intI) = excluc(0, intI) + excluc(1, intI)
            exclue(2, intI) = exclue(0, intI) + exclue(1, intI)
            notapc(2, intI) = notapc(0, intI) + notapc(1, intI)
            notape(2, intI) = notape(0, intI) + notape(1, intI)
            smissc(2, intI) = smissc(0, intI) + smissc(1, intI)
            smisse(2, intI) = smisse(0, intI) + smisse(1, intI)
            validc(2, intI) = validc(0, intI) + validc(1, intI)
            valide(2, intI) = valide(0, intI) + valide(1, intI)
            If validc(2, intI) > 0 Then
               resltc(2, intI) = (validc(0, intI) * resltc(0, intI) _
                  + validc(1, intI) * resltc(1, intI)) / validc(2, intI)
            Else
               resltc(2, intI) = -1.0
            End If
            If valide(2, intI) > 0 Then
               reslte(2, intI) = (valide(0, intI) * reslte(0, intI) _
                  + valide(1, intI) * reslte(1, intI)) / valide(2, intI)
            Else
               reslte(2, intI) = -1.0
            End If
         Next
         nsqct(2) = nsqct(0) + nsqct(1)
         nsqet(2) = nsqet(0) + nsqet(1)
         excluct(2) = excluct(0) + excluct(1)
         excluet(2) = excluet(0) + excluet(1)
         notapct(2) = notapct(0) + notapct(1)
         notapet(2) = notapet(0) + notapet(1)
         smissct(2) = smissct(0) + smissct(1)
         smisset(2) = smisset(0) + smisset(1)
         validct(2) = validct(0) + validct(1)
         validet(2) = validet(0) + validet(1)
         ' Note that the use of CInt below gives a more accurate result,
         ' for integer endpoint runs, because the numerator always has to be
         ' an integer; it represents the total number of males and females 
         ' with a positive endpoint.
         If validct(2) > 0 Then
            If bolFloatPt Then
               resltct(2) = (validct(0) * resltct(0) _
                  + validct(1) * resltct(1)) / validct(2)
            Else
               resltct(2) = CDbl(CInt(validct(0) * resltct(0) _
                  + validct(1) * resltct(1))) / validct(2)
            End If
         Else
            resltct(2) = -1.0
         End If
         If validet(2) > 0 Then
            If bolFloatPt Then
               resltet(2) = (validet(0) * resltet(0) _
                  + validet(1) * resltet(1)) / validet(2)
            Else
               resltet(2) = CDbl(CInt(validet(0) * resltet(0) _
                  + validet(1) * resltet(1))) / validet(2)
            End If
         Else
            resltet(2) = -1.0
         End If
      End If
   End Sub

   Private Sub PermPval(ByRef p1 As Double, ByRef p2 As Double, ByVal deltaH0 As Double)
      '------------------------------------------------------------------
      ' OVERVIEW
      ' This procedure performs a permutation test based on a paired
      ' randomized assignment of treatment (control vs experimental) to 
      ' experimental units.  (See Edgington, Encycl of Stat. Sciences, 
      ' Vol 7, pp 530-538, for a description of permutation tests).

      ' The permutation test for the HSYAS randomized trial proceeds as 
      ' follows:  

      ' 1. Denote by D a statistic that measures the difference in
      '    treatment effect on some endpoint variable.  Denote by DOBS the
      '    observed value from our data.  It is defined in terms of
      '    (control-experimental) prevalence differences, so *negative*
      '    values are "good" for endpoints like 6-month cessation.
      ' 2. Then consider 2^M permutations of the data, where each
      '    permutation corresponds to one of the possible randomized 
      '    assignments.  For each permutation there is a corresponding 
      '    value of the statistic D.  (The actual observed statistic DOBS
      '    is one of these values, corresponding to the permutation that
      '    gives the actual randomization assignment).
      ' 3. Under the null hypothesis Ho: no treatment effect, each of the
      '    2^M permutations is equally likely and hence the 2^M different
      '    values of D are equally likely.
      ' 4. The proportion p of the 2^M values of D that are greater than or
      '    equal to the observed value DOBS is the attained one-sided p-value
      '    (because under Ho Pr(D >= DOBS)=p).  The two-sided p-value is
      '    obtained from this via a simple computation.

      ' DESCRIPTION OF VARIABLES
      '    m - number of pairs of schools with the outcome variable
      '    nlt,neq,ngt - running count of number of permutations for which
      '                  the statistics are less than, equal to or greater
      '                  than the reference value dobs.
      '    imc(i) - # of individuals in control school i, i=1,..,m, 
      '    / dmc(i)   with outcome variable=1, if endpoint is binary.
      '               In general, it is the sum of values for the endpoint.
      '    inc(i) - total # of individuals in control school i, i=1,..,m
      '    ime(i) - # of individuals in experimental school i, i=1,..,m,  
      '    / dme(i)   with outcome variable=1, if endpoint is binary.
      '               In general, it is the sum of values for the endpoint.
      '    ine(i) - total # of individuals in experimental school i, i=1,..,m 
      '    d(i) -  difference in prevalences between the control and
      '            experimental schools in pair i, i=1,...,m
      '    deltaH0 Null hypothesis value for the underlying difference
      '            between control and experimental prevalences
      '    dobs -  permutation test statistic for the unpermuted data
      '            under delta = deltaH0
      '    dobs0 - permutation test statistic for the unpermuted data
      '            under delta = 0.0
      '    dperm - permutation test statistic for the current permutation
      '    p1, p2 - one-sided and two-sided p-values for dobs
      '------------------------------------------------------------------
      Dim dobs As Double, dperm As Double, pleft As Double, ptotmc As Double
      Dim ptotme As Double, ptotnc As Double, ptotne As Double, pmcsave As Double
      Dim pncsave As Double, pmc(24) As Double, pme(24) As Double
      Dim pnc(24) As Double, pne(24) As Double, ind(24) As Integer
      Dim nlt As Integer, neq As Integer, ngt As Integer
      Dim i As Integer, j As Integer, k As Integer

      ' Compute sample sizes and prevalence differences.  Initialize a
      ' flag ind(i), i = 1,..,m, for each cluster, which is used to
      ' keep track of permutations.  Pmc, pme, pnc, and pne are
      ' initialized to the unpermuted counts in imc, ime, inc and ine.
      ' Also do this for the corresponding total count variables.
      For i = 0 To m - 1
         ind(i) = 0
         If bolFloatPt Then
            pmc(i) = dmc(i)
            pme(i) = dme(i)
         Else
            pmc(i) = imc(i)
            pme(i) = ime(i)
         End If
         pnc(i) = inc(i)
         pne(i) = ine(i)
      Next
      If bolFloatPt Then
         ptotmc = dtotmc
         ptotme = dtotme
      Else
         ptotmc = totmc
         ptotme = totme
      End If
      ptotnc = totnc
      ptotne = totne

      ' dobs holds the test statistic for testing H0: delta = deltaH0.
      dobs = dobs0 - deltaH0

      ' The test is equivalent to subtracting deltaH0 from the observed
      ' prevalence differences, then testing H0: delta = 0.  This is in
      ' turn equivalent to adding deltaH0 to the observed control
      ' prevalences.
      For i = 0 To m - 1
         pmc(i) = pmc(i) + pnc(i) * deltaH0
      Next
      ptotmc = ptotmc + ptotnc * deltaH0

      ' Initialize counts of permutations.
      nlt = 0
      neq = 1
      ngt = 0

      ' Enter permutation loop
      Do
         ' Initialize school pair to first pair in each iteration of main loop
         i = 0

         ' If flag is set proceed to next pair
         While i < m AndAlso ind(i) = 1
            i += 1
         End While

         ' When i=m, all 2^m permutations have been examined
         If i < m Then
            ' Set flag & unset all lower flags (to cycle through lower permutations)
            ind(i) = 1
            j = i - 1
            ' Note that the loop statement is skipped if j=-1
            For k = 0 To j
               ind(k) = 0
            Next

            ' Flip data to create a new permutation.  This changes the
            ' totals as follows:
            ptotmc = ptotmc + pme(i) - pmc(i)
            ptotme = ptotme + pmc(i) - pme(i)
            ptotnc = ptotnc + pne(i) - pnc(i)
            ptotne = ptotne + pnc(i) - pne(i)
            pmcsave = pmc(i)
            pncsave = pnc(i)
            pmc(i) = pme(i)
            pnc(i) = pne(i)
            pme(i) = pmcsave
            pne(i) = pncsave

            ' Compute difference statistic for this permutation
            dperm = ptotme / ptotne - ptotmc / ptotnc

            ' Tally permutations with dperm<dobs, dperm=dobs, or dperm>dobs
            ' The floating point version takes a more conservative approach
            ' to assessing "equality" of permutation statistics, in order to
            ' allow for any rounding error.
            If bolFloatPt Then
               If Math.Abs(dperm - dobs) < 0.0000001 Then
                  neq += 1
               ElseIf dperm - dobs > 0.0000001 Then
                  ngt += 1
               Else
                  nlt += 1
               End If
            Else
               If dperm < dobs Then
                  nlt += 1
               ElseIf dperm > dobs Then
                  ngt += 1
               Else
                  neq += 1
               End If
            End If
         End If
      Loop Until i = m

      ' Calculate one-sided p-value
      p1 = CDbl(neq + ngt) / CDbl(nlt + neq + ngt)

      ' Calculate two-sided p-value
      pleft = CDbl(neq + nlt) / CDbl(nlt + neq + ngt)
      If pleft <= p1 Then
         p2 = 2.0 * pleft
      Else
         p2 = 2.0 * p1
      End If
      If p2 > 1.0 Then p2 = 1.0

   End Sub

   Private Sub PermCI(ByRef dl As Double, ByRef du As Double)
      ' This subroutine computes a 95% confidence interval (dl,du) for
      ' the weighted difference in prevalence (control - experimental).
      ' See Tom Braun's 1999 thesis, pp 42-49 and Gail and Mark,
      ' Statistics in Medicine, vol. 15, 1069-1092 (1996) for references

      ' Set constants use in convergence criteria
      Const eps As Double = 0.000025
      Const tol As Double = 0.001

      ' Set t(df,.975) for df=1, ..., 24
      Dim t975() As Double = {12.7062047361747, 4.30265272974946, 3.18244630528356, _
         2.77644510519779, 2.570581835615, 2.44691185114486, 2.36462425159278, _
         2.30600413519914, 2.2621571627982, 2.22813885198627, 2.20098516009164, _
         2.1788128296672298, 2.16036865646279, 2.14478668791779, 2.13144954555975, _
         2.11990529922122, 2.10981557783327, 2.10092204024098, 2.09302405440824, _
         2.08596344726578, 2.07961384472758, 2.07387306790391, 2.06865761041892, _
         2.06389856162789}

      Dim dmean As Double, ssquare As Double, dhalfint As Double
      Dim dl0 As Double, du0 As Double, d1 As Double, d2 As Double
      Dim dold As Double, dnew As Double, p2d1 As Double, p2d2 As Double
      Dim p2dold As Double, p2dnew As Double, r As Double
      Dim iter As Integer

      ' If m<6, there is no CI, because the minimum value of the p-value 
      ' function is 2^(-m+1).  Set flags for infinity in this case.

      If m < 6 Then
         dl = -1000000.0
         du = +1000000.0
         Exit Sub
      End If

      ' Compute unweighted mean prevalence difference

      dmean = 0.0
      For i As Integer = 0 To m - 1
         dmean = dmean + d(i)
      Next
      dmean = dmean / (CDbl(m))

      ' Compute starting values for dl and du

      ssquare = 0.0
      For i As Integer = 0 To m - 1
         ssquare = ssquare + (d(i) - dmean) ^ 2
      Next
      ssquare = ssquare / CDbl(m - 1)
      dl0 = dmean - Math.Sqrt(ssquare / CDbl(m)) * t975(m - 2)
      du0 = dmean + Math.Sqrt(ssquare / CDbl(m)) * t975(m - 2)

      Dim p1 As Double
      PermPval(p1, p2d1, dobs0)

      ' Start lower bound calculations.  Begin with an interval (d1,d2)
      ' that contains dl0 (in most cases dl0 is one endpoint of the
      ' interval); compute p-values for each endpoint and then
      ' use linear interpolation to produce a new estimate.

      iter = 0

      If dl0 < 0 Then
         d1 = dl0
         d2 = dl0 * 0.75
      ElseIf dl0 > 0 Then
         d1 = dl0
         d2 = dl0 * 1.25
      Else
         d1 = -0.05
         d2 = 0.05
      End If

      ' In ill-behaved cases, d1 may be greater than dobs0; use a
      ' different start interval if this is true.

      If d1 > dobs0 Then
         d2 = dobs0
         If dobs0 <> 0 Then
            d1 = dobs0 - Math.Abs(dobs0) / 2.0
         Else
            d1 = -0.05
         End If
      End If

      pval2lb(p2d1, d1)
      pval2lb(p2d2, d2)

      dold = d1
      p2dold = p2d1

      ' Loop: interpolate to get a new estimate

      Do
         dhalfint = Math.Abs(d2 - d1) / 2.0

         ' The following is needed in case the initial interval is "too
         ' small", with endpoints that have the same p-value.

         While p2d1 = p2d2
            If d2 < dobs0 Then
               d2 = Math.Min(d2 + dhalfint, dobs0)
               pval2lb(p2d2, d2)
            Else
               d1 = d1 - dhalfint
               pval2lb(p2d1, d1)
            End If
         End While

         r = (p2d2 - 0.049975) / (p2d2 - p2d1)
         dnew = r * d1 + (1.0 - r) * d2
         iter += 1

         ' Get p-value corresponding to dnew

         pval2lb(p2dnew, dnew)

         ' If p2dnew is close enough to 0.049975, accept dnew as lower bound.

         If Math.Abs(p2dnew - 0.049975) < eps Then
            dl = dnew
            Exit Do
         End If

         ' The discreteness of the p-value function (especially for small
         ' numbers of clusters) requires the added conditions to prevent
         ' an infinite loop when succeeding iterates have p-values on either
         ' side of 0.049975 which differ by the minimum amount (2^(-m+1))
         ' or when the iterates are "close" together but their p-values
         ' do not satisfy the convergence criterion.  In these situations,
         ' apply a simple binary search to obtain a good confidence interval.

         If (p2dnew < 0.049975 And p2dold > 0.049975 Or p2dnew > 0.049975 And p2dold < 0.049975) _
               And (Math.Abs(p2dnew - p2dold) = 2.0 ^ (-m + 1) Or Math.Abs(dold - dnew) < tol) Then
            txtProgress.AppendText(vbCrLf & _
               "Switching to binary search for UB in iteration " & iter.ToString)
            txtProgress.ScrollToCaret()
            ' Force screen updates to appear *now*
            Application.DoEvents()
            If dold < dnew Then
               BinSrch(dold, dnew, p2dold, dl, 1)
            Else
               BinSrch(dnew, dold, p2dnew, dl, 1)
            End If
            Exit Do
         End If

         ' Otherwise, adjust interval (d1,d2) and iterate

         If p2dnew <= p2d1 Then
            d1 = dnew
            p2d1 = p2dnew
         Else
            If p2dnew > 0.049975 Then
               d2 = dnew
               p2d2 = p2dnew
            Else
               d1 = dnew
               p2d1 = p2dnew
            End If
         End If
         dold = dnew
         p2dold = p2dnew
      Loop

      ' Use a similar procedure to get an upper bound

      iter = 0

      If du0 < 0 Then
         d1 = du0 * 1.25
         d2 = du0
      ElseIf du0 > 0 Then
         d1 = du0 * 0.75
         d2 = du0
      Else
         d1 = -0.05
         d2 = 0.05
      End If

      ' In ill-behaved cases, d2 may be less than dobs0; use a
      ' different start interval if this is true.

      If d2 < dobs0 Then
         d1 = dobs0
         If dobs0 <> 0 Then
            d2 = dobs0 + Math.Abs(dobs0) / 2.0
         Else
            d2 = 0.05
         End If
      End If

      pval2ub(p2d1, d1)
      pval2ub(p2d2, d2)

      dold = d2
      p2dold = p2d2

      ' Loop: interpolate to get a new estimate

      Do
         dhalfint = Math.Abs(d2 - d1) / 2.0

         ' The following is needed in case the initial interval is "too
         ' small", with endpoints that have the same p-value.

         While p2d1 = p2d2
            If d1 > dobs0 Then
               d1 = Math.Max(d1 - dhalfint, dobs0)
               pval2ub(p2d1, d1)
            Else
               d2 = d2 + dhalfint
               pval2ub(p2d2, d2)
            End If
         End While

         r = (p2d2 - 0.950025) / (p2d2 - p2d1)
         dnew = r * d1 + (1.0 - r) * d2
         iter += 1

         ' Get p-value corresponding to dnew

         pval2ub(p2dnew, dnew)

         ' If 1-p2dnew is close enough to 0.049975, accept dnew as upper
         ' bound.

         If Math.Abs(p2dnew - 0.950025) < eps Then
            du = dnew
            Exit Do
         End If

         ' Apply a test for the infinite loop condition that is analogous 
         ' to the one used for the lower bound.

         If (p2dnew < 0.950025 And p2dold > 0.950025 Or p2dnew > 0.950025 And p2dold < 0.950025) _
            And (Math.Abs(p2dnew - p2dold) = 2.0 ^ (-m + 1) Or Math.Abs(dold - dnew) < tol) Then
            txtProgress.AppendText(vbCrLf & _
               "Switching to binary search for UB in iteration " & iter.ToString)
            txtProgress.ScrollToCaret()
            ' Force screen updates to appear *now*
            Application.DoEvents()
            If dold < dnew Then
               BinSrch(dold, dnew, p2dnew, du, 0)
            Else
               BinSrch(dnew, dold, p2dold, du, 0)
            End If
            Exit Do
         End If

         ' Otherwise, adjust interval (d1,d2) and iterate

         If p2dnew >= p2d2 Then
            d2 = dnew
            p2d2 = p2dnew
         Else
            If p2dnew > 0.950025 Then
               d2 = dnew
               p2d2 = p2dnew
            Else
               d1 = dnew
               p2d1 = p2dnew
            End If
         End If
         dold = dnew
         p2dold = p2dnew
      Loop

   End Sub

   Private Sub BinSrch(ByVal d1 As Double, ByVal d2 As Double, ByVal pval As Double, _
      ByRef dresult As Double, ByVal lb As Integer)
      ' This subroutine uses a binary search to find a delta with p(delta)=pval 
      ' that is within 1.0E-05 of the point where the p-value function 
      ' changes to the next higher value (if lb=1) or the next lower value 
      ' (if lb=0).  This is intended for use in situations where the p-value 
      ' function is "too discrete", so that the normal convergence criterion 
      ' has no solution.
      Const Tolerance As Double = 0.00001
      Dim da As Double, db As Double, dmid As Double, pdmid As Double

      ' The starting interval (d1,d2) contains the point where the p-value 
      ' function makes the discrete jump of interest.  This procedure 
      ' successively halves this interval, choosing the half that spans 
      ' the jump point at each iteration.  This continues until the length
      ' of the interval is less than Tolerance.  The endpoint that is chosen 
      ' as the result is the one that is on the high side for the upper 
      ' CI bound and on the low side for the lower CI bound.  This 
      ' will insure that the CI is always at least of level 95%.

      da = d1
      db = d2

      Do
         dmid = (da + db) / 2.0
         If lb = 1 Then
            pval2lb(pdmid, dmid)
         Else
            pval2ub(pdmid, dmid)
         End If
         If pdmid = pval Then
            If lb = 1 Then
               da = dmid
            Else
               db = dmid
            End If
            dresult = dmid
         Else
            If lb = 1 Then
               db = dmid
               dresult = da
            Else
               da = dmid
               dresult = db
            End If
         End If
      Loop Until Math.Abs(da - db) < Tolerance

   End Sub

   Private Sub pval2lb(ByRef pval As Double, ByVal delta As Double)
      ' This subroutine computes a truncated version of the two-sided
      ' p-value funtion, used for finding the lower endpoint of the
      ' confidence interval.  This is a nondecreasing function of delta.
      Dim p1 As Double
      If delta < dobs0 Then
         PermPval(p1, pval, delta)
      Else
         pval = 1.0
      End If
   End Sub

   Private Sub pval2ub(ByRef pval As Double, ByVal delta As Double)
      ' This subroutine computes a truncated version of (1 - two-sided
      ' p-value funtion), used for finding the upper endpoint of the
      ' confidence interval.  This is a nondecreasing function of delta.
      Dim p1 As Double
      If delta > dobs0 Then
         PermPval(p1, pval, delta)
         pval = 1.0 - pval
      Else
         pval = 0.0
      End If
   End Sub

   Private Sub detail_out(ByRef docDetail As Word.Document, ByVal Gender As String, _
      ByRef pageno As Integer, ByVal m As Integer, _
      ByRef usepair() As Boolean, ByVal r1pc As Double, ByVal dobs As Double, _
      ByVal p2 As Double, ByVal dl As Double, ByVal du As Double)
      ' Subroutine creates a Word document for detailed output for the current run
      ' and uses Word methods to append formatted paragraphs to it.
      Dim strPval2 As String, strR1PC As String, strLastChar As String
      Dim strDU As String, strDL As String, strR1DL As String, strR1DU As String
      Dim inmark As String, intI As Integer, includedc As Integer, includede As Integer
      Dim missc As Integer, misse As Integer, delta As Double, includedct As Integer
      Dim includedet As Integer, missct As Integer, misset As Integer, lngI As Long
      Dim nsqt As Integer, validt As Integer, validratio As Single
      Dim r1dl As Double, r1du As Double, bolMainEndPt As Boolean, intC As Integer
      Dim objPara As Word.Paragraph
      Dim objRange As Word.Range
      Dim objLine As Word.Shape

      bolMainEndPt = (mainendp = 1)
      pageno += 1
      ' Insert a page break before adding text to pages 2, 3, etc.
      ' This is done here, rather than at the end of formatting for the 
      ' previous page, to prevent a spurious blank page after the last page.
      If pageno > 1 Then
         objRange = docDetail.Content.Paragraphs.Last.Range
         objRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
         objRange.InsertBreak(Word.WdBreakType.wdPageBreak)
      End If
      ' Clear any existing tab stops.
      objRange = docDetail.Bookmarks.Item("\endofdoc").Range
      objRange.ParagraphFormat.TabStops.ClearAll()
      ' Insert the first paragraph.  The line spacing is set to 11 pt in
      ' the top section to help make everything fit on one page
      objPara = docDetail.Content.Paragraphs.Last
      objPara.Range.Font.Bold = False
      objPara.Range.Text = "HSYAS Analyses Using Cluster-Size " & _
         "Weighted Permutation Tests without Covariates" & vbCrLf
      With objRange.ParagraphFormat
         .LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly
         .LineSpacing = 11
      End With
      If pageno > 1 Then
         ' Strip out the extra paragraph mark that Word insists on adding
         ' after the page break.  Note that the deletion has to occur directly
         ' on the document object, not the range object!  Hoever, objRange.End
         ' returns the character position in terms of the entire document text.
         objRange = docDetail.Paragraphs(docDetail.Content.Paragraphs.Count - 2).Range
         strLastChar = Strings.Right(objRange.Text, 1)
         If InStr(strLastChar, Chr(13)) > 0 Then
            intC = objRange.End
            docDetail.Range(intC - 1, intC).Delete()
         End If
      End If
      objRange = docDetail.Bookmarks.Item("\endofdoc").Range
      objRange.InsertParagraphAfter()

      objRange = docDetail.Bookmarks.Item("\endofdoc").Range
      objRange.InsertAfter("Group")
      objRange.Font.Underline = Word.WdUnderline.wdUnderlineSingle
      objRange = docDetail.Bookmarks.Item("\endofdoc").Range
      objRange.InsertAfter(": " & Trim(group) & " " & Gender)
      objRange.Font.Underline = Word.WdUnderline.wdUnderlineNone
      objRange.ParagraphFormat.TabStops.Add(5.54 * 72, Word.WdTabAlignment.wdAlignTabLeft)
      objRange.ParagraphFormat.TabStops.Add(6.39 * 72, Word.WdTabAlignment.wdAlignTabLeft)
      objRange.InsertAfter(vbTab & Format(DateTime.Now, "d") & _
         vbTab & "Page " & pageno & vbCrLf)
      objRange.InsertParagraphAfter()

      objRange = docDetail.Bookmarks.Item("\endofdoc").Range
      objRange.ParagraphFormat.TabStops.ClearAll()
      objRange.InsertAfter("Subgroup")
      objRange.Font.Underline = Word.WdUnderline.wdUnderlineSingle
      objRange = docDetail.Bookmarks.Item("\endofdoc").Range
      objRange.InsertAfter(": " & subgroup & vbCrLf)
      objRange.Font.Underline = Word.WdUnderline.wdUnderlineNone
      objRange.InsertParagraphAfter()

      objRange = docDetail.Bookmarks.Item("\endofdoc").Range
      objRange.InsertAfter("Category: " & category & vbCrLf)
      objRange.InsertParagraphAfter()

      objRange = docDetail.Bookmarks.Item("\endofdoc").Range
      objRange.InsertAfter("Endpoint number " & varnum)
      If bolFloatPt Then
         objRange.InsertAfter(" (FP)")
      End If
      objRange.Font.Bold = bolMainEndPt
      objRange.InsertParagraphAfter()

      objRange = docDetail.Bookmarks.Item("\endofdoc").Range
      objRange.ParagraphFormat.TabStops.Add(0.25 * 72, Word.WdTabAlignment.wdAlignTabLeft)
      objRange.InsertAfter(vbTab & varname)
      objRange.Font.Bold = bolMainEndPt
      objRange.InsertParagraphAfter()

      objRange = docDetail.Bookmarks.Item("\endofdoc").Range
      objRange.InsertAfter(vbTab & vardescr & vbCrLf)
      objRange.Font.Bold = bolMainEndPt
      objRange.InsertParagraphAfter()

      objRange = docDetail.Bookmarks.Item("\endofdoc").Range
      objRange.ParagraphFormat.TabStops.ClearAll()
      objRange.InsertAfter("Pairs marked with ""**"" were included in the " & _
         "permutation test." & vbCrLf)
      objRange.InsertParagraphAfter()

      objRange = docDetail.Bookmarks.Item("\endofdoc").Range
      objRange.ParagraphFormat.TabStops.Add(1.63 * 72, Word.WdTabAlignment.wdAlignTabLeft)
      objRange.ParagraphFormat.TabStops.Add(4.51 * 72, Word.WdTabAlignment.wdAlignTabLeft)
      objRange.InsertAfter(vbTab & "CONTROLS" & vbTab & "EXPERIMENTALS" & vbCrLf)
      objLine = docDetail.Shapes.AddLine(42.9, 6, 110.9, 6, objRange)
      objLine.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
      objLine.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph
      objLine.Line.Weight = 3
      objLine.Line.Style = Microsoft.Office.Core.MsoLineStyle.msoLineThinThin
      objLine.SetShapesDefaultProperties()
      objLine = docDetail.Shapes.AddLine(184.2, 6, 252.2, 6, objRange)
      objLine = docDetail.Shapes.AddLine(267.8, 6, 319.8, 6, objRange)
      objLine = docDetail.Shapes.AddLine(422.5, 6, 474.5, 6, objRange)
      objRange.InsertParagraphAfter()

      objRange = docDetail.Bookmarks.Item("\endofdoc").Range
      objRange.ParagraphFormat.TabStops.ClearAll()
      objRange.ParagraphFormat.TabStops.Add(0.58 * 72, Word.WdTabAlignment.wdAlignTabLeft)
      objRange.ParagraphFormat.TabStops.Add(1.93 * 72, Word.WdTabAlignment.wdAlignTabLeft)
      objRange.ParagraphFormat.TabStops.Add(3.76 * 72, Word.WdTabAlignment.wdAlignTabLeft)
      objRange.ParagraphFormat.TabStops.Add(5.09 * 72, Word.WdTabAlignment.wdAlignTabLeft)
      objRange.InsertAfter(vbTab & "Surveys recvd/" & vbTab & "Missing/" & _
         vbTab & "Surveys recvd/" & vbTab & "Missing/")
      objRange.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle
      objRange.InsertParagraphAfter()

      objRange = docDetail.Bookmarks.Item("\endofdoc").Range
      objRange.ParagraphFormat.TabStops.Add(3.16 * 72, Word.WdTabAlignment.wdAlignTabCenter)
      objRange.ParagraphFormat.TabStops.Add(6.28 * 72, Word.WdTabAlignment.wdAlignTabCenter)
      objRange.ParagraphFormat.TabStops.Add(6.97 * 72, Word.WdTabAlignment.wdAlignTabCenter)
      objRange.InsertAfter("School" & vbTab & "Number with" & vbTab & "NA/ SQ" & _
         vbTab & "Endpoint" & vbTab & "Number with" & vbTab & "NA/ SQ" & _
         vbTab & "Endpoint" & vbTab & "Delta")
      objRange.ParagraphFormat.LeftIndent = -5.75
      objRange.InsertParagraphAfter()

      objRange = docDetail.Bookmarks.Item("\endofdoc").Range
      objRange.ParagraphFormat.TabStops.Add(0.02 * 72, Word.WdTabAlignment.wdAlignTabLeft)
      objRange.InsertAfter(vbTab & "Pair" & vbTab & "useable data" & vbTab & "Excluded" & _
         vbTab & "Result" & vbTab & "useable data" & vbTab & "Excluded" & _
         vbTab & "Result" & vbTab & "(E-C)")
      objRange.InsertParagraphAfter()

      objRange = docDetail.Bookmarks.Item("\endofdoc").Range
      objRange.InsertAfter(vbCrLf)
      objLine = docDetail.Shapes.AddLine(-9.1, 7.15, 31.2, 7.15, objRange)
      objLine.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
      objLine.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph
      objLine.Line.Weight = 1
      objLine.Line.Style = Microsoft.Office.Core.MsoLineStyle.msoLineSingle
      objLine.SetShapesDefaultProperties()
      objLine = docDetail.Shapes.AddLine(42.9, 7.15, 123.5, 7.15, objRange)
      objLine = docDetail.Shapes.AddLine(139.1, 7.15, 189.8, 7.15, objRange)
      objLine = docDetail.Shapes.AddLine(202.8, 7.15, 253.5, 7.15, objRange)
      objLine = docDetail.Shapes.AddLine(271.7, 7.15, 352.3, 7.15, objRange)
      objLine = docDetail.Shapes.AddLine(366.6, 7.15, 417.3, 7.15, objRange)
      objLine = docDetail.Shapes.AddLine(427.7, 7.15, 478.4, 7.15, objRange)
      objLine = docDetail.Shapes.AddLine(486.2, 7.15, 518.7, 7.15, objRange)

      For intI = 0 To 24
         If usepair(intI) Then
            inmark = "**"
         Else
            inmark = ""
         End If
         ' smissc and smisse have been left in the code, even though they
         ' will almost certainly be zero in all HSYAS runs.  This will make
         ' for less work when adapting HSPP SPSS programs for setting up
         ' the input datasets.  Thus, includedc and includede can be teated
         ' as the total number of surveys received in each arm of the study.
         includedc = nsqc(sex - 1, intI) - smissc(sex - 1, intI)
         includede = nsqe(sex - 1, intI) - smisse(sex - 1, intI)
         missc = includedc - validc(sex - 1, intI) - excluc(sex - 1, intI) - notapc(sex - 1, intI)
         misse = includede - valide(sex - 1, intI) - exclue(sex - 1, intI) - notape(sex - 1, intI)
         If resltc(sex - 1, intI) <> -1.0 And reslte(sex - 1, intI) <> -1.0 Then
            delta = reslte(sex - 1, intI) - resltc(sex - 1, intI)
         Else
            delta = 0.0
         End If
         ' Some paragraph formatting only needs to be set once, for the first line
         If intI = 0 Then
            objRange = docDetail.Bookmarks.Item("\endofdoc").Range
            objRange.ParagraphFormat.TabStops.ClearAll()
            objRange.ParagraphFormat.TabStops.Add(0.29 * 72, Word.WdTabAlignment.wdAlignTabRight)
            objRange.ParagraphFormat.TabStops.Add(0.92 * 72, Word.WdTabAlignment.wdAlignTabRight)
            objRange.ParagraphFormat.TabStops.Add(1.1 * 72, Word.WdTabAlignment.wdAlignTabRight)
            objRange.ParagraphFormat.TabStops.Add(1.44 * 72, Word.WdTabAlignment.wdAlignTabRight)
            objRange.ParagraphFormat.TabStops.Add(2.13 * 72, Word.WdTabAlignment.wdAlignTabRight)
            objRange.ParagraphFormat.TabStops.Add(2.46 * 72, Word.WdTabAlignment.wdAlignTabRight)
            objRange.ParagraphFormat.TabStops.Add(2.64 * 72, Word.WdTabAlignment.wdAlignTabRight)
            objRange.ParagraphFormat.TabStops.Add(3.03 * 72, Word.WdTabAlignment.wdAlignTabDecimal)
            objRange.ParagraphFormat.TabStops.Add(4.1 * 72, Word.WdTabAlignment.wdAlignTabRight)
            objRange.ParagraphFormat.TabStops.Add(4.28 * 72, Word.WdTabAlignment.wdAlignTabRight)
            objRange.ParagraphFormat.TabStops.Add(4.62 * 72, Word.WdTabAlignment.wdAlignTabRight)
            objRange.ParagraphFormat.TabStops.Add(5.31 * 72, Word.WdTabAlignment.wdAlignTabRight)
            objRange.ParagraphFormat.TabStops.Add(5.63 * 72, Word.WdTabAlignment.wdAlignTabRight)
            objRange.ParagraphFormat.TabStops.Add(5.8 * 72, Word.WdTabAlignment.wdAlignTabRight)
            objRange.ParagraphFormat.TabStops.Add(6.12 * 72, Word.WdTabAlignment.wdAlignTabDecimal)
            objRange.ParagraphFormat.TabStops.Add(6.79 * 72, Word.WdTabAlignment.wdAlignTabDecimal)
            objRange.ParagraphFormat.LeftIndent = -9.35
         End If
         objRange.InsertAfter(inmark & vbTab & (intI + 1) & vbTab & _
            nsqc(sex - 1, intI) & vbTab & "/" & vbTab & validc(sex - 1, intI) & vbTab & _
            missc & "/" & vbTab & notapc(sex - 1, intI) & "/" & vbTab & _
            excluc(sex - 1, intI) & vbTab & Format(resltc(sex - 1, intI), "###.0000") & _
            vbTab & nsqe(sex - 1, intI) & vbTab & "/" & vbTab & valide(sex - 1, intI) & _
            vbTab & misse & "/" & vbTab & notape(sex - 1, intI) & "/" & vbTab & _
            exclue(sex - 1, intI) & vbTab & Format(reslte(sex - 1, intI), "###.0000") & _
            vbTab & Format(delta, "###.0000") & vbCrLf)
      Next

      ' Add in a row of separators prior to the overall totals
      objRange = docDetail.Bookmarks.Item("\endofdoc").Range
      objRange.InsertAfter(vbCrLf)
      objLine = docDetail.Shapes.AddLine(49.2, 8.1, 109.0, 8.1, objRange)
      objLine = docDetail.Shapes.AddLine(135.2, 8.1, 197.6, 8.1, objRange)
      objLine = docDetail.Shapes.AddLine(209.3, 8.1, 253.5, 8.1, objRange)
      objLine = docDetail.Shapes.AddLine(276, 8.1, 338.4, 8.1, objRange)
      objLine = docDetail.Shapes.AddLine(361.4, 8.1, 425.1, 8.1, objRange)
      objLine = docDetail.Shapes.AddLine(435.5, 8.1, 471.9, 8.1, objRange)
      objLine = docDetail.Shapes.AddLine(482.3, 8.1, 518.7, 8.1, objRange)

      includedct = nsqct(sex - 1) - smissct(sex - 1)
      includedet = nsqet(sex - 1) - smisset(sex - 1)
      missct = includedct - validct(sex - 1) - excluct(sex - 1) - notapct(sex - 1)
      misset = includedet - validet(sex - 1) - excluet(sex - 1) - notapet(sex - 1)
      If resltct(sex - 1) <> -1.0 And resltet(sex - 1) <> -1.0 Then
         delta = resltet(sex - 1) - resltct(sex - 1)
      Else
         delta = 0.0
      End If

      ' Prepare tab stops for the Overall totals row and write it
      objRange = docDetail.Bookmarks.Item("\endofdoc").Range
      objRange.ParagraphFormat.TabStops.ClearAll()
      objRange.ParagraphFormat.TabStops.Add(1.01 * 72, Word.WdTabAlignment.wdAlignTabRight)
      objRange.ParagraphFormat.TabStops.Add(1.12 * 72, Word.WdTabAlignment.wdAlignTabRight)
      objRange.ParagraphFormat.TabStops.Add(1.52 * 72, Word.WdTabAlignment.wdAlignTabRight)
      objRange.ParagraphFormat.TabStops.Add(2.17 * 72, Word.WdTabAlignment.wdAlignTabRight)
      objRange.ParagraphFormat.TabStops.Add(2.49 * 72, Word.WdTabAlignment.wdAlignTabRight)
      objRange.ParagraphFormat.TabStops.Add(2.76 * 72, Word.WdTabAlignment.wdAlignTabRight)
      objRange.ParagraphFormat.TabStops.Add(3.03 * 72, Word.WdTabAlignment.wdAlignTabDecimal)
      objRange.ParagraphFormat.TabStops.Add(4.19 * 72, Word.WdTabAlignment.wdAlignTabRight)
      objRange.ParagraphFormat.TabStops.Add(4.3 * 72, Word.WdTabAlignment.wdAlignTabRight)
      objRange.ParagraphFormat.TabStops.Add(4.69 * 72, Word.WdTabAlignment.wdAlignTabRight)
      objRange.ParagraphFormat.TabStops.Add(5.34 * 72, Word.WdTabAlignment.wdAlignTabRight)
      objRange.ParagraphFormat.TabStops.Add(5.67 * 72, Word.WdTabAlignment.wdAlignTabRight)
      objRange.ParagraphFormat.TabStops.Add(5.9 * 72, Word.WdTabAlignment.wdAlignTabRight)
      objRange.ParagraphFormat.TabStops.Add(6.12 * 72, Word.WdTabAlignment.wdAlignTabDecimal)
      objRange.ParagraphFormat.TabStops.Add(6.79 * 72, Word.WdTabAlignment.wdAlignTabDecimal)
      objRange.ParagraphFormat.LeftIndent = 0
      objRange.InsertAfter("Overall:" & vbTab & nsqct(sex - 1) & vbTab & "/" & _
         vbTab & validct(sex - 1) & vbTab & missct & "/" & vbTab & _
         notapct(sex - 1) & "/" & vbTab & excluct(sex - 1) & vbTab & _
         Format(resltct(sex - 1), "###.0000") & vbTab & nsqet(sex - 1) & _
         vbTab & "/" & vbTab & validet(sex - 1) & vbTab & misset & "/" & _
         vbTab & notapet(sex - 1) & "/" & vbTab & excluet(sex - 1) & vbTab & _
         Format(resltet(sex - 1), "###.0000") & vbTab & _
         Format(delta, "###.0000") & vbCrLf)
      objRange.InsertParagraphAfter()
      objRange = docDetail.Bookmarks.Item("\endofdoc").Range
      objRange.ParagraphFormat.TabStops.ClearAll()

      nsqt = nsqct(sex - 1) + nsqet(sex - 1)
      validt = validct(sex - 1) + validet(sex - 1)
      'ratio1 = CSng(includedt) / CSng(nsqct(sex - 1) + nsqet(sex - 1))
      If nsqt > 0 Then
         validratio = CSng(validt) / CSng(nsqt)
      Else
         validratio = 0.0
      End If

      objPara = docDetail.Content.Paragraphs.Add
      objPara.Range.Text = "% of surveys with useable data for this item (" & _
         validt & " of " & nsqt & "): " & Format(validratio, "##0.00%") & vbCrLf
      objPara.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly
      objPara.LineSpacing = 11
      objRange = docDetail.Bookmarks.Item("\endofdoc").Range
      objRange.InsertParagraphAfter()

      objRange = docDetail.Bookmarks.Item("\endofdoc").Range
      objRange.InsertAfter("Number of school pairs providing data for " & _
         "the permutation test: " & m & vbCrLf)
      objRange.InsertParagraphAfter()

      If m = 0 Then
         objRange = docDetail.Bookmarks.Item("\endofdoc").Range
         objRange.InsertAfter("THERE IS INSUFFICIENT DATA FOR THE PERMUTATION TEST")
         objRange.Font.Bold = True
         objRange.InsertParagraphAfter()
      Else
         If p2 > 0 And p2 < 0.0001 Then
            strPval2 = "<0.0001"
         Else
            strPval2 = p2.ToString("#.0000")
         End If
         If r1pc = -1000000.0 Then
            strR1PC = " Not Avail."
            strR1DL = " Not Avail."
            strR1DU = " Not Avail."
         Else
            strR1PC = r1pc.ToString("F4") + "%"
            If dl = -1000000.0 Then
               strR1DL = " Not Avail."
            Else
               r1dl = dl / resltct(sex - 1) * 100.0
               strR1DL = r1dl.ToString("F4") + "%"
            End If
            If du = +1000000.0 Then
               strR1DU = " Not Avail."
            Else
               r1du = du / resltct(sex - 1) * 100.0
               strR1DU = r1du.ToString("F4") + "%"
            End If
         End If
         If dl = -1000000.0 Then
            strDL = " -infinity"
         Else
            strDL = dl.ToString("###.0000000")
         End If
         If du = +1000000.0 Then
            strDU = " +infinity"
         Else
            strDU = du.ToString("###.0000000")
         End If

         objRange = docDetail.Bookmarks.Item("\endofdoc").Range
         objRange.ParagraphFormat.TabStops.Add(3.41 * 72, Word.WdTabAlignment.wdAlignTabDecimal)
         objRange.InsertAfter("Permutation test statistic and 95% CI:" & vbTab & _
            Format(dobs, "###.0000000") & " ( " & strDL & ", " & strDU & ")" & vbCrLf)
         objRange.InsertParagraphAfter()
         objRange = docDetail.Bookmarks.Item("\endofdoc").Range
         objRange.InsertAfter("Permutation test p-value (2-sided):" & vbTab & _
            strPval2 & vbCrLf)
         objRange.InsertParagraphAfter()
         objRange = docDetail.Bookmarks.Item("\endofdoc").Range
         objRange.InsertAfter("Percent increase and 95% CI:" & vbTab & _
            strR1PC & " ( " & strR1DL & ", " & strR1DU & ")")
      End If
      ' Atempt to reformat any paragraph that has acquired "after" spacing of 10 points.  There
      ' doesn't seem to be any other way to combat this other than looping through every
      ' paragraph and testing for aberrant formatting.
      For lngI = 1 To docDetail.Paragraphs.Count
         objRange = docDetail.Paragraphs(lngI).Range
         If objRange.ParagraphFormat.SpaceAfter = 10 Then
            objRange.ParagraphFormat.SpaceAfter = 0
            If objRange.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple Then
               objRange.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly
               objRange.ParagraphFormat.LineSpacing = 11
            End If
         End If
      Next
   End Sub

   Private Sub summary_out(ByRef docSummary As Word.Document, ByVal Gender As String, _
      ByRef spageno As Integer, ByRef slineno As Integer, _
      ByVal m As Integer, ByVal r1pc As Double, ByVal dobs As Double, _
      ByVal p2 As Double, ByVal dl As Double, ByVal du As Double)
      ' Subroutine summarizes the current run in one line and appends to docSummary
      Dim strPval2 As String, strR1PC As String, validt As Integer
      Dim strDObs As String, strDU As String, strDL As String, bolMainEndPt As Boolean
      Dim PagePos As Single
      Dim objRange As Word.Range
      Dim objLine As Word.Shape

      bolMainEndPt = (mainendp = 1)

      ' Handle special cases where some output variables are undefined.
      ' Decimal points are inserted in some of the special messages because
      ' the Word output uses decimal tabs for positioning; if the decimal points
      ' were omitted, the result would be less readable.
      If m = 0 Then
         strPval2 = "NO.DATA"
         strR1PC = "NO.DATA"
         strDObs = "NO.DATA"
         strDU = "NO DATA"
         strDL = "NO DATA"
      Else
         strDObs = dobs.ToString("F4")
         If dl = -1000000.0 Then
            strDL = "-infin"
         Else
            strDL = dl.ToString("F4")
         End If
         If du = 1000000.0 Then
            strDU = "+infin"
         Else
            strDU = du.ToString("F4")
         End If
         If r1pc = -1000000.0 Then
            strR1PC = ".N/A "
         Else
            strR1PC = r1pc.ToString("F2") + "%"
         End If
         ' Temporarily display higher precision for test purposes
         'strPval2 = p2.ToString("0.0000000000E-00")
         If p2 > 0 And p2 < 0.0001 Then
            strPval2 = "<0.0001"
         Else
            strPval2 = p2.ToString("F4")
         End If
      End If

      ' Write a page heading if this is a new page.  The determination of
      ' whether this is a new page is based on whether the pageno is zero
      ' (start of output) or if the position where the next text will appear
      ' is less than a line height below the top margin.  The value that
      ' Word returns for the first line on the page is not exactly at the
      ' top margin, but is slightly below it.
      objRange = docSummary.Bookmarks.Item("\endofdoc").Range
      PagePos = objRange.Information(Word.WdInformation.wdVerticalPositionRelativeToPage)

      If PagePos < docSummary.PageSetup.TopMargin + objRange.ParagraphFormat.LineSpacing Then
         spageno += 1
         objRange = docSummary.Bookmarks.Item("\endofdoc").Range
         objRange.ParagraphFormat.TabStops.ClearAll()
         objRange.ParagraphFormat.TabStops.Add(7.24 * 72, Word.WdTabAlignment.wdAlignTabLeft)
         objRange.ParagraphFormat.TabStops.Add(8.25 * 72, Word.WdTabAlignment.wdAlignTabLeft)
         objRange.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly
         objRange.ParagraphFormat.LineSpacing = 11.5
         objRange.InsertAfter("HSYAS Analyses Using Cluster-Size " & _
            "Weighted Permutation Tests without Covariates" & vbTab & _
            Format(DateTime.Now, "d") & vbTab & "Page " & spageno & vbCrLf)
         objRange = docSummary.Bookmarks.Item("\endofdoc").Range
         objRange.InsertParagraphAfter()

         objRange = docSummary.Bookmarks.Item("\endofdoc").Range
         objRange.ParagraphFormat.TabStops.ClearAll()
         objRange.ParagraphFormat.TabStops.Add(3.18 * 72, Word.WdTabAlignment.wdAlignTabLeft)
         objRange.InsertAfter("Summary of Results for " & group)
         objRange.Font.Underline = Word.WdUnderline.wdUnderlineSingle
         objRange = docSummary.Bookmarks.Item("\endofdoc").Range
         objRange.InsertAfter(vbTab)
         objRange.Font.Underline = Word.WdUnderline.wdUnderlineNone
         objRange = docSummary.Bookmarks.Item("\endofdoc").Range
         objRange.InsertAfter("Subgroup")
         objRange.Font.Underline = Word.WdUnderline.wdUnderlineSingle
         objRange = docSummary.Bookmarks.Item("\endofdoc").Range
         objRange.InsertAfter(": " & subgroup & vbCrLf)
         objRange.Font.Underline = Word.WdUnderline.wdUnderlineNone
         objRange.InsertParagraphAfter()

         objRange = docSummary.Bookmarks.Item("\endofdoc").Range
         objRange.ParagraphFormat.TabStops.ClearAll()
         objRange.ParagraphFormat.TabStops.Add(1.5 * 72, Word.WdTabAlignment.wdAlignTabCenter)
         objRange.ParagraphFormat.TabStops.Add(2.83 * 72, Word.WdTabAlignment.wdAlignTabLeft)
         objRange.ParagraphFormat.TabStops.Add(4.13 * 72, Word.WdTabAlignment.wdAlignTabCenter)
         objRange.ParagraphFormat.TabStops.Add(4.91 * 72, Word.WdTabAlignment.wdAlignTabLeft)
         objRange.ParagraphFormat.TabStops.Add(5.51 * 72, Word.WdTabAlignment.wdAlignTabLeft)
         objRange.ParagraphFormat.TabStops.Add(7.26 * 72, Word.WdTabAlignment.wdAlignTabLeft)
         objRange.ParagraphFormat.TabStops.Add(8.25 * 72, Word.WdTabAlignment.wdAlignTabLeft)
         objRange.InsertAfter(vbTab & "Useable" & vbTab & "Control" & vbTab & _
            "Experimental" & vbTab & "Delta" & vbTab & "95% Confidence" & vbTab & "Percent" & _
            vbTab & "2-sided")
         objRange.InsertParagraphAfter()

         objRange = docSummary.Bookmarks.Item("\endofdoc").Range
         objRange.ParagraphFormat.TabStops.Add(0.16 * 72, Word.WdTabAlignment.wdAlignTabLeft)
         objRange.ParagraphFormat.TabStops.Add(1.93 * 72, Word.WdTabAlignment.wdAlignTabLeft)
         objRange.InsertAfter(vbTab & "Endpoints" & vbTab & "Data" & vbTab & "Gender" & _
            vbTab & "Results" & vbTab & "Results" & vbTab & "(E-C)" & _
            vbTab & "Interval for Delta" & vbTab & "Increase" & vbTab & "p-value")
         objRange.InsertParagraphAfter()

         objRange = docSummary.Bookmarks.Item("\endofdoc").Range
         objRange.InsertAfter(vbCrLf)
         objLine = docSummary.Shapes.AddLine(0, 3.8, 75.4, 3.8, objRange)
         objLine.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
         objLine.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph
         objLine.Line.Weight = 1
         objLine.Line.Style = Microsoft.Office.Core.MsoLineStyle.msoLineSingle
         objLine.SetShapesDefaultProperties()
         objLine = docSummary.Shapes.AddLine(84.5, 3.8, 131.3, 3.8, objRange)
         objLine = docSummary.Shapes.AddLine(139.8, 3.8, 189.8, 3.8, objRange)
         objLine = docSummary.Shapes.AddLine(202.8, 3.8, 248.3, 3.8, objRange)
         objLine = docSummary.Shapes.AddLine(261.3, 3.8, 332.8, 3.8, objRange)
         objLine = docSummary.Shapes.AddLine(344.5, 3.8, 388.7, 3.8, objRange)
         objLine = docSummary.Shapes.AddLine(397.8, 3.8, 487.8, 3.8, objRange)
         objLine = docSummary.Shapes.AddLine(521.4, 3.8, 573.25, 3.8, objRange)
         objLine = docSummary.Shapes.AddLine(594.1, 3.8, 632.25, 3.8, objRange)

         objRange = docSummary.Bookmarks.Item("\endofdoc").Range
         objRange.ParagraphFormat.TabStops.ClearAll()
         objRange.ParagraphFormat.TabStops.Add(0.22 * 72, Word.WdTabAlignment.wdAlignTabRight)
         objRange.ParagraphFormat.TabStops.Add(0.27 * 72, Word.WdTabAlignment.wdAlignTabLeft)
         objRange.ParagraphFormat.TabStops.Add(1.66 * 72, Word.WdTabAlignment.wdAlignTabRight)
         objRange.ParagraphFormat.TabStops.Add(1.93 * 72, Word.WdTabAlignment.wdAlignTabLeft)
         objRange.ParagraphFormat.TabStops.Add(2.98 * 72, Word.WdTabAlignment.wdAlignTabDecimal)
         objRange.ParagraphFormat.TabStops.Add(3.99 * 72, Word.WdTabAlignment.wdAlignTabDecimal)
         objRange.ParagraphFormat.TabStops.Add(4.97 * 72, Word.WdTabAlignment.wdAlignTabDecimal)
         objRange.ParagraphFormat.TabStops.Add(5.51 * 72, Word.WdTabAlignment.wdAlignTabLeft)
         objRange.ParagraphFormat.TabStops.Add(7.53 * 72, Word.WdTabAlignment.wdAlignTabDecimal)
         objRange.ParagraphFormat.TabStops.Add(8.41 * 72, Word.WdTabAlignment.wdAlignTabDecimal)
      End If
      ' Write the one-line summary for the current run
      objRange = docSummary.Bookmarks.Item("\endofdoc").Range
      objRange.InsertAfter(vbTab & varnum & ":" & vbTab & varname)
      objRange.Font.Bold = bolMainEndPt
      objRange = docSummary.Bookmarks.Item("\endofdoc").Range
      validt = validct(sex - 1) + validet(sex - 1)
      objRange.InsertAfter(vbTab & validt & vbTab & Gender & vbTab & _
         Format(resltct(sex - 1), "###.0000") & vbTab & _
         Format(resltet(sex - 1), "###.0000") & vbTab & _
         strDObs & vbTab & "(" & strDL & ", " & strDU & ")" & vbTab & _
         strR1PC & vbTab & strPval2 & vbCrLf)
      objRange.Font.Bold = False
   End Sub

   Private Sub mnuExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuExit.Click
      Me.Close()
   End Sub

   Private Sub frmPermTest_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
      ' Before closing, check to see if a hidden copy of Word was created 
      ' and is still open.  If so, close it.
      If Not bolWordAlreadyOpen And Not (appWord Is Nothing) Then
         appWord.Quit(Word.WdSaveOptions.wdDoNotSaveChanges)
         appWord = Nothing
      End If
   End Sub

   Private Sub mnuAbout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuAbout.Click
      Dim AboutForm As New frmAbout
      AboutForm.ShowDialog(Me)
   End Sub
End Class
