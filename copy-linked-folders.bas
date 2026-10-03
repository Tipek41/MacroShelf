Attribute VB_Name = "CopyLinkedFolders"
Option Explicit

Sub LinkliKlasorleriKopyala()

    Dim secim As Range, hucre As Range
    Dim fso As Object, gorulen As Object, secici As Object
    Dim kaynak As String, hedef As String, yeni As String
    Dim ad As String, anaHedef As String
    Dim kopya As Long, atlanan As Long, hata As Long, n As Long
    Dim hataNo As Long, hataMetni As String, detay As String

    On Error Resume Next
    Set secim = Application.InputBox( _
        Prompt:="Kopyalanacak klasor linklerinin bulundugu hucreleri secin." & _
                vbCrLf & "Ayri alanlar icin Ctrl tusunu kullanabilirsiniz.", _
        Title:="Linkleri sec", Type:=8)
    On Error GoTo GenelHata

    If secim Is Nothing Then Exit Sub

    Set secici = Application.FileDialog(4)
    With secici
        .Title = "Klasorlerin kopyalanacagi yeri secin"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        hedef = .SelectedItems(1)
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set gorulen = CreateObject("Scripting.Dictionary")
    gorulen.CompareMode = vbTextCompare

    For Each hucre In secim.Cells

        kaynak = KlasorLinkiniOku(hucre)

        If Len(kaynak) = 0 Then
            atlanan = atlanan + 1
            GoTo Sonraki
        End If

        If Left$(kaynak, 2) <> "\\" And Mid$(kaynak, 2, 1) <> ":" Then
            If Len(hucre.Parent.Parent.Path) = 0 Then
                hata = hata + 1
                GoTo Sonraki
            End If
            kaynak = fso.BuildPath(hucre.Parent.Parent.Path, kaynak)
        End If

        kaynak = fso.GetAbsolutePathName(kaynak)

        If gorulen.Exists(kaynak) Then
            atlanan = atlanan + 1
            GoTo Sonraki
        End If
        gorulen.Add kaynak, True

        If Not fso.FolderExists(kaynak) Then
            hata = hata + 1
            If hata <= 8 Then detay = detay & vbCrLf & _
                hucre.Address(False, False) & ": Klasor bulunamadi / erisim yok."
            GoTo Sonraki
        End If

        anaHedef = fso.GetAbsolutePathName(hedef)
        If LCase$(anaHedef) = LCase$(kaynak) Or _
           InStr(1, anaHedef & "\", kaynak & "\", vbTextCompare) = 1 Then
            hata = hata + 1
            If hata <= 8 Then detay = detay & vbCrLf & _
                hucre.Address(False, False) & ": Hedef kaynak klasorun icinde."
            GoTo Sonraki
        End If

        ad = fso.GetFolder(kaynak).Name
        yeni = fso.BuildPath(hedef, ad)
        n = 1

        Do While fso.FolderExists(yeni) Or fso.FileExists(yeni)
            n = n + 1
            yeni = fso.BuildPath(hedef, ad & "_" & n)
        Loop

        Application.StatusBar = "Kopyalaniyor: " & ad
        DoEvents

        On Error Resume Next
        Err.Clear
        fso.CopyFolder kaynak, yeni, False
        hataNo = Err.Number
        hataMetni = Err.Description
        On Error GoTo GenelHata

        If hataNo = 0 Then
            kopya = kopya + 1
        Else
            hata = hata + 1
            If hata <= 8 Then detay = detay & vbCrLf & _
                ad & ": " & hataMetni
        End If

Sonraki:
    Next hucre

    Application.StatusBar = False

    MsgBox "Kopyalanan klasor: " & kopya & vbCrLf & _
           "Atlanan bos / tekrar hucre: " & atlanan & vbCrLf & _
           "Basarisiz klasor: " & hata & vbCrLf & detay & vbCrLf & vbCrLf & _
           "Hedef: " & hedef & _
           IIf(hata > 0, vbCrLf & _
           "Kopyalama hatasi alan klasorler hedefte eksik kalmis olabilir.", ""), _
           vbInformation, "Kopyalama sonucu"
    Exit Sub

GenelHata:
    Application.StatusBar = False
    MsgBox "Islem durdu: " & Err.Description & vbCrLf & _
           "Tamamlanan klasor: " & kopya, vbExclamation

End Sub

Private Function KlasorLinkiniOku(ByVal hucre As Range) As String

    Dim yol As String
    Dim re As Object, eslesme As Object

    On Error GoTo Okunamadi

    If hucre.Hyperlinks.Count > 0 Then
        yol = hucre.Hyperlinks(1).Address
    End If

    If Len(yol) = 0 And hucre.HasFormula Then
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "^\s*=\s*HYPERLINK\s*\(\s*(""([^""]|"""")*""|[^,;]+)"
        re.IgnoreCase = True

        If re.Test(hucre.Formula) Then
            Set eslesme = re.Execute(hucre.Formula)(0)
            yol = CStr(hucre.Parent.Evaluate(eslesme.SubMatches(0)))
        End If
    End If

    If Len(yol) = 0 And Not IsError(hucre.Value2) Then
        yol = CStr(hucre.Value2)
    End If

    yol = Trim$(yol)

    If LCase$(Left$(yol, 8)) = "file:///" Then
        yol = Mid$(yol, 9)
    ElseIf LCase$(Left$(yol, 7)) = "file://" Then
        yol = "\\" & Mid$(yol, 8)
    End If

    yol = Replace(yol, "%20", " ")
    yol = Replace(yol, "/", "\")

    If InStr(1, yol, "://", vbTextCompare) > 0 Then Exit Function

    Do While Right$(yol, 1) = "\" And Len(yol) > 3
        yol = Left$(yol, Len(yol) - 1)
    Loop

    KlasorLinkiniOku = yol
    Exit Function

Okunamadi:
    KlasorLinkiniOku = ""

End Function
