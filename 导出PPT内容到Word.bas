Attribute VB_Name = "����PPT���ݵ�Word"
Sub ����PPT���ݵ�Word()
    Dim ppPres As Presentation '��ǰPPT
    Dim objSlide As Object '�õ�Ƭ����
    Dim objshapes As Object 'shape����
    Dim shap As Object 'shape����
    Dim path_fd_putput As String
    Dim i_Slide As Integer '���ڱ����õ�Ƭҳ��
    Dim i_shap As Integer '���ڱ���ÿһҳǶ���ĵ���
    Dim imgPath As String
    path_fd_putput = "C:\Users\Geoyee\Desktop\output" '����ļ���
    Set ppPres = ActivePresentation
    For i_Slide = 1 To ppPres.Slides.Count '�����õ�Ƭ
        Set objSlide = ppPres.Slides(i_Slide)
        i_shap = 0
        For Each shap In objSlide.Shapes '�����õ�Ƭ�е�shape����
            If shap.Type = 7 Then  '��Ƕ����
                If shap.OLEFormat.ProgID = "Word.Document.8" Or _
                   shap.OLEFormat.ProgID = "Word.Document.12" Then '�ж��ĵ������Ƿ�Ϊword
                    i_shap = i_shap + 1
                    Set objshapes = shap.OLEFormat.Object
                    objshapes.SaveAs2 FileName:=path_fd_putput & "\" & i_Slide & "_" & i_shap & "_ole.docx", _
                        FileFormat:=12, LockComments:=False, Password:="", _
                        AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
                        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
                        :=False, SaveAsAOCELetter:=False, CompatibilityMode:=15
                        Set objshapes = Nothing
                End If
            ElseIf shap.Type = 17 Or shap.Type = 1 Then  '�ı���
                i_shap = i_shap + 1
                Dim wordDoc As New Word.Document
                Set woedDoc = Word.Application.Documents.Add
                wordDoc.Range().Text = wordDoc.Range().Text + shap.TextFrame.TextRange.Text
                wordDoc.SaveAs2 FileName:=path_fd_putput & "\" & i_Slide & "_" & i_shap & "_txt.docx", _
                    FileFormat:=12, LockComments:=False, Password:="", _
                    AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
                    EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
                    :=False, SaveAsAOCELetter:=False, CompatibilityMode:=15
                wordDoc.Close
                Set wordDoc = Nothing
            ElseIf shap.Type = 13 Then  'ͼƬ
                i_shap = i_shap + 1
                Dim wordImg As New Word.Document
                Set wordImg = Word.Application.Documents.Add
                imgPath = path_fd_putput & "\" & i_Slide & "_" & i_shap & "_tmp.png"
                shap.Export imgPath, ppShapeFormatPNG
                wordImg.InlineShapes.AddPicture FileName:=imgPath, SaveWithDocument:=True
                wordImg.SaveAs2 FileName:=path_fd_putput & "\" & i_Slide & "_" & i_shap & "_img.docx", FileFormat:=12
                wordImg.Close
                Set wordImg = Nothing
                If FileExists(imgPath) Then
                    Kill imgPath
                End If
            End If
        Next shap
    Next i_Slide
    Set ppPres = Nothing
End Sub


Function FileExists(ByVal FileToTest As String)
   FileExists = (Dir(FileToTest) <> "")
End Function
