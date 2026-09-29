Attribute VB_Name = "PdfFolderMerger"
Option Explicit

' Merges top-level PDF files in a selected folder, sorted by filename.
' Requires Microsoft Excel for Windows and Adobe Acrobat desktop (not Reader).
Public Sub MergePDFsInFolder()
    Dim folderPath As String, folderName As String
    Dim outputPath As String, tempOutputPath As String
    Dim fso As Object, folder As Object, file As Object
    Dim pdfPaths() As String, pdfCount As Long, i As Long, j As Long
    Dim swapPath As String, validCount As Long, failedCount As Long
    Dim acroApp As Object, mergedDoc As Object, sourceDoc As Object
    Dim pages As Long, insertOK As Boolean, answer As VbMsgBoxResult
    Dim completed As Boolean, errorText As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the folder containing the PDFs to merge"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    folderName = folder.Name
    outputPath = fso.BuildPath(folderPath, folderName & ".pdf")
    tempOutputPath = fso.BuildPath(folderPath, folderName & "_merging_temp.pdf")

    If fso.FileExists(outputPath) Then
        answer = MsgBox("The output already exists:" & vbCrLf & outputPath & vbCrLf & vbCrLf & _
                        "Replace it after a successful merge?", vbQuestion + vbYesNo, "PDF Merger")
        If answer <> vbYes Then Exit Sub
    End If

    For Each file In folder.Files
        If LCase$(fso.GetExtensionName(file.Name)) = "pdf" Then
            If StrComp(file.Path, outputPath, vbTextCompare) <> 0 And _
               StrComp(file.Path, tempOutputPath, vbTextCompare) <> 0 Then
                pdfCount = pdfCount + 1
                ReDim Preserve pdfPaths(1 To pdfCount)
                pdfPaths(pdfCount) = file.Path
            End If
        End If
    Next file

    If pdfCount = 0 Then
        MsgBox "No source PDF files were found in the selected folder.", vbExclamation, "PDF Merger"
        Exit Sub
    End If

    ' Sort full paths alphabetically (case-insensitive).
    For i = 1 To pdfCount - 1
        For j = i + 1 To pdfCount
            If StrComp(fso.GetFileName(pdfPaths(i)), fso.GetFileName(pdfPaths(j)), vbTextCompare) > 0 Then
                swapPath = pdfPaths(i)
                pdfPaths(i) = pdfPaths(j)
                pdfPaths(j) = swapPath
            End If
        Next j
    Next i

    On Error GoTo MergeError
    Set acroApp = CreateObject("AcroExch.App")
    Set mergedDoc = CreateObject("AcroExch.PDDoc")
    If Not mergedDoc.Create Then Err.Raise vbObjectError + 100, , "Could not create the output PDF."

    For i = 1 To pdfCount
        Set sourceDoc = CreateObject("AcroExch.PDDoc")
        If sourceDoc.Open(pdfPaths(i)) Then
            pages = sourceDoc.GetNumPages
            If pages > 0 Then
                insertOK = mergedDoc.InsertPages(mergedDoc.GetNumPages - 1, sourceDoc, 0, pages, False)
                If insertOK Then
                    validCount = validCount + 1
                Else
                    failedCount = failedCount + 1
                End If
            Else
                failedCount = failedCount + 1
            End If
            sourceDoc.Close
        Else
            failedCount = failedCount + 1
        End If
        Set sourceDoc = Nothing
    Next i

    If validCount = 0 Or mergedDoc.GetNumPages = 0 Then
        Err.Raise vbObjectError + 101, , "None of the source PDFs could be merged. Check that they are not damaged or password-protected."
    End If

    If fso.FileExists(tempOutputPath) Then fso.DeleteFile tempOutputPath, True
    If Not mergedDoc.Save(1, tempOutputPath) Then Err.Raise vbObjectError + 102, , "Could not save the temporary output PDF."

    mergedDoc.Close
    Set mergedDoc = Nothing
    acroApp.Exit
    Set acroApp = Nothing

    ' Replace the previous output only after the new PDF was created successfully.
    fso.CopyFile tempOutputPath, outputPath, True
    fso.DeleteFile tempOutputPath, True
    completed = True
    On Error GoTo 0

    If failedCount = 0 Then
        MsgBox "PDFs merged successfully." & vbCrLf & validCount & " file(s) included." & vbCrLf & outputPath, vbInformation, "PDF Merger"
    Else
        MsgBox "Merge completed with some files skipped." & vbCrLf & _
               validCount & " file(s) included; " & failedCount & " file(s) could not be read." & vbCrLf & outputPath, vbExclamation, "PDF Merger"
    End If
    Exit Sub

MergeError:
    errorText = Err.Description
    On Error Resume Next
    If Not sourceDoc Is Nothing Then sourceDoc.Close
    If Not mergedDoc Is Nothing Then mergedDoc.Close
    If Not acroApp Is Nothing Then acroApp.Exit
    If fso.FileExists(tempOutputPath) Then fso.DeleteFile tempOutputPath, True
    On Error GoTo 0
    MsgBox "The merge could not be completed:" & vbCrLf & errorText & vbCrLf & vbCrLf & _
           "Check that Adobe Acrobat desktop is installed and that the PDFs are accessible.", vbCritical, "PDF Merger"
End Sub
