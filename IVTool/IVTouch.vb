Module IVTouch
    Private ivApp As Inventor.Application
    Private pid As Integer
    'Private ivApp As Inventor.ApprenticeServerComponent
    Private Const BaseVaultPath = "C:\VaultWorkspace"
    Private Const BaseDWFSavePath = "C:\Projects"

    Public Sub LoadIV()
        Dim ivProcesses() As Process

        ivApp = CreateObject("Inventor.Application")
        ivApp.Visible = False
        ivApp.SilentOperation = True
        ivApp.ScreenUpdating = False
        ivProcesses = Process.GetProcessesByName("Inventor")
        pid = ivProcesses(ivProcesses.GetLowerBound(0)).Id
    End Sub

    Public Sub CloseIV()
        Process.GetProcessById(pid).Kill()
    End Sub

    Public Sub ProcessAssembly(ByVal docName As String, ByVal bRec As Boolean, ByVal bTB As Boolean)
        Dim filePath As String

        filePath = RecursiveFind(BaseVaultPath, docName + ".iam")
        If filePath <> "" Then
            frm_IVTool.Cout("Found document")
            frm_IVTool.Cout("+ " + My.Computer.FileSystem.GetName(filePath))

            'ivApp = New Inventor.ApprenticeServerComponent
            'ivApp.ApplicationAddIns.ItemById("{0AC6FD95-2F4D-42CE-8BE0-8AEA580399E4}").Activate()
            ProcessDrawing(filePath, bRec, bTB)

        Else
            frm_IVTool.Cout("No document found")
        End If
    End Sub

    Private Sub ProcessDrawing(ByVal docPath As String, ByVal bRec As Boolean, ByVal bTB As Boolean)
        Dim drawingPath As String
        Dim oDoc As Inventor.Document

        drawingPath = FindDrawing(docPath)
        If drawingPath <> "" Then
            frm_IVTool.Cout("Found drawing")
            frm_IVTool.Cout("++ " + My.Computer.FileSystem.GetName(drawingPath))
            oDoc = ivApp.Documents.Open(drawingPath, False)
            'PublishDWF(oDoc, My.Computer.FileSystem.GetName(drawingPath))
            oDoc.Close()
        End If

    End Sub

    Private Function FindDrawing(docPath As String) As String
        Dim baseDocName As String

        baseDocName = My.Computer.FileSystem.GetName(docPath)
        baseDocName = Left(baseDocName, Len(baseDocName) - 4)

        Return RecursiveFind(BaseVaultPath, baseDocName + ".idw")

    End Function

    Private Sub ReplaceTitleblock(docPath As String)

    End Sub

    Private Function RecursiveFind(ByVal folderPath As String, ByVal fileName As String) As String
        Dim filePath As String

        filePath = folderPath + "\" + fileName
        RecursiveFind = ""

        If My.Computer.FileSystem.FileExists(filePath) Then
            RecursiveFind = filePath
        Else
            For Each foundDir As String In My.Computer.FileSystem.GetDirectories(folderPath)
                RecursiveFind = RecursiveFind + RecursiveFind(foundDir, fileName)
            Next
        End If
    End Function

    Private Sub PublishDWF(ByRef oDoc As Inventor.Document, ByVal dwgNum As String)
        ' Get the DWF translator Add-In.
        'Dim DWFAddIn As Inventor.TranslatorAddIn
        Dim DWFAddIn As Object
        DWFAddIn = DirectCast(ivApp.ApplicationAddIns.ItemById("{0AC6FD95-2F4D-42CE-8BE0-8AEA580399E4}"), Inventor.TranslatorAddIn)

        'Set a reference to the active document (the document to be published).
        'Dim odocument As Document
        'Set odocument = ThisApplication.ActiveDocument

        'Copy model iproperties to drawing iproperties
        Dim propSetDU As Inventor.PropertySet
        propSetDU = oDoc.PropertySets("Inventor User Defined Properties")

        Dim propSetDS As Inventor.PropertySet
        propSetDS = oDoc.PropertySets("Inventor Summary Information")

        Dim propSetMU As Inventor.PropertySet
        propSetMU = oDoc.Sheets.Item(1).DrawingViews.Item(1).ReferencedDocumentDescriptor.ReferencedDocument.PropertySets.Item("Inventor User Defined Properties")

        Dim propSetMS As Inventor.PropertySet
        propSetMS = oDoc.Sheets.Item(1).DrawingViews.Item(1).ReferencedDocumentDescriptor.ReferencedDocument.PropertySets.Item("Inventor Summary Information")

        If Exists(propSetMU, "SL12_Process Code") Then ' eg: "01"
            If Exists(propSetDU, "SL12_Process Code") Then
                propSetDU.Item("SL12_Process Code").Value = propSetMU.Item("SL12_Process Code").Value
            Else
                Call propSetDU.Add(propSetMU.Item("SL12_Process Code").Value, "SL12_Process Code")
            End If
        End If

        If Exists(propSetMU, "SL02_Product_Code") Then ' eg: "D-10-SLICKLINE UNITS"
            If Exists(propSetDU, "SL02_Product_Code") Then
                propSetDU.Item("SL02_Product_Code").Value = propSetMU.Item("SL02_Product_Code").Value
            Else
                Call propSetDU.Add(propSetMU.Item("SL02_Product_Code").Value, "SL02_Product_Code")
            End If
        End If

        If Exists(propSetMU, "SL13_Product Category") Then ' eg: "D13"
            If Exists(propSetDU, "SL13_Product Category") Then
                propSetDU.Item("SL13_Product Category").Value = propSetMU.Item("SL13_Product Category").Value
            Else
                Call propSetDU.Add(propSetMU.Item("SL13_Product Category").Value, "SL13_Product Category")
            End If
        End If

        If Exists(propSetMU, "SL04_Drawing_Document_Type") Then ' eg: "Principal Assembly"
            If Exists(propSetDU, "SL04_Drawing_Document_Type") Then
                propSetDU.Item("SL04_Drawing_Document_Type").Value = propSetMU.Item("SL04_Drawing_Document_Type").Value
            Else
                Call propSetDU.Add(propSetMU.Item("SL04_Drawing_Document_Type").Value, "SL04_Drawing_Document_Type")
            End If
        End If

        If Exists(propSetMS, "Title") Then ' eg: "SPOOLER"
            If Exists(propSetDS, "Title") Then
                propSetDS.Item("Title").Value = propSetMS.Item("Title").Value
            Else
                Call propSetDS.Add(propSetMS.Item("Title").Value, "Title")
            End If
        End If

        Dim oContext As Inventor.TranslationContext
        oContext = ivApp.TransientObjects.CreateTranslationContext
        oContext.Type = Inventor.IOMechanismEnum.kFileBrowseIOMechanism

        ' Create a NameValueMap object
        Dim oOptions As Inventor.NameValueMap
        oOptions = ivApp.TransientObjects.CreateNameValueMap

        ' Create a DataMedium object
        Dim oDataMedium As Inventor.DataMedium
        oDataMedium = ivApp.TransientObjects.CreateDataMedium

        Dim voption As Object

        ' Check whether the translator has 'SaveCopyAs' options
        If DWFAddIn.HasSaveCopyAsOptions(oDoc, oContext, oOptions) Then


            oOptions.Value("Launch_Viewer") = 0

            ' Other options...
            oOptions.Value("Publish_All_Component_Props") = 0
            oOptions.Value("Publish_All_Physical_Props") = 0
            'oOptions.Value("Password") = 0

            If TypeOf oDoc Is Inventor.DrawingDocument Then

                ' Drawing options
                oOptions.Value("Publish_Mode") = Inventor.DWFPublishModeEnum.kCustomDWFPublish
                oOptions.Value("Publish_All_Sheets") = 1
            End If

        End If

        'Set the destination file name
        oDataMedium.FileName = BaseDWFSavePath + "\" + dwgNum + ".dwf"
        'oDataMedium.FileName = fSavePath + ".dwf"

        'Publish document.
        Call DWFAddIn.SaveCopyAs(oDoc, oContext, oOptions, oDataMedium)


    End Sub

    Private Function Exists(ByVal oCol As Inventor.PropertySet, ByVal vKey As String) As Boolean
        Dim o As Object
        Try
            o = oCol.Item(vKey)
            Exists = True
        Catch ex As Exception
            Exists = ExistsNonObject(oCol, vKey)
        End Try
    End Function

    Private Function ExistsNonObject(ByVal oCol As Inventor.PropertySet, ByVal vKey As String) As Boolean
        Dim v As Object
        Try
            v = oCol.Item(vKey)
            ExistsNonObject = True
        Catch ex As Exception
            ExistsNonObject = False
        End Try
    End Function

End Module
