Attribute VB_Name = "导出PPT内容到Word"


Sub 导出PPT内容到Word()
    Dim ppPres As Presentation '当前PPT
    Dim objSlide As Object '幻灯片对象
    Dim objshapes As Object 'shape对象
    Dim shap As Object 'shape对象
    Dim path_fd_putput As String
    Dim i_Slide As Integer '用于遍历幻灯片页数
    Dim i_shap As Integer '用于遍历每一页嵌入文档数
    Dim imgPath As String
    path_fd_putput = "C:\Users\Geoyee\Desktop\output" '输出文件夹
    Set ppPres = ActivePresentation
    For i_Slide = 1 To ppPres.Slides.Count '遍历幻灯片
        Set objSlide = ppPres.Slides(i_Slide)
        i_shap = 0
        For Each shap In objSlide.Shapes '遍历幻灯片中的shape对象
            If shap.Type = 7 Then  '内嵌对象
                If shap.OLEFormat.ProgID = "Word.Document.8" Or _
                   shap.OLEFormat.ProgID = "Word.Document.12" Then '判断文档类型是否为word
                    i_shap = i_shap + 1
                    Set objshapes = shap.OLEFormat.Object
                    objshapes.SaveAs2 FileName:=path_fd_putput & "\" & i_Slide & "_" & i_shap & "_ole.docx", _
                        FileFormat:=12, LockComments:=False, Password:="", _
                        AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
                        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
                        :=False, SaveAsAOCELetter:=False, CompatibilityMode:=15
                        Set objshapes = Nothing
                End If
            ElseIf shap.Type = 17 Or shap.Type = 1 Then  '文本框
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
            ElseIf shap.Type = 13 Then  '图片
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
