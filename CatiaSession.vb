Option Explicit On
Option Strict On

Public Class CatiaSession

    Private ReadOnly _app As INFITF.Application
    Private ReadOnly _status As CatiaSessionStatus

    Public Sub New()
        _app = Connect()
        _status = EvaluateStatus(_app)
    End Sub

    Private Function Connect() As INFITF.Application
        Try
            Return CType(GetObject(, "CATIA.Application"), INFITF.Application)
        Catch
            Return Nothing
        End Try
    End Function



    Private Function EvaluateStatus(app As INFITF.Application) As CatiaSessionStatus
        If app Is Nothing Then Return CatiaSessionStatus.NotRunning
        If app.Windows.Count = 0 Then Return CatiaSessionStatus.NoWindowsOpen

        Dim oActiveDoc As INFITF.Document = app.ActiveDocument
        Dim typeNameDoc As String = TypeName(oActiveDoc)

        ' 1. Validación base: ¿El documento activo tiene ruta y está guardado?
        Dim activeSaved As Boolean = Not String.IsNullOrEmpty(oActiveDoc.Path) AndAlso oActiveDoc.Saved
        If Not activeSaved Then
            If typeNameDoc = "ProductDocument" Then Return CatiaSessionStatus.ProductDocumentNotSaved
            If typeNameDoc = "DrawingDocument" Then Return CatiaSessionStatus.DrawingDocumentNotSaved
            Return CatiaSessionStatus.Unknown
        End If

        ' 2. Validación de integridad (Documentos relacionados)
        ' Si es Product o Drawing, revisamos que TODO lo que esté cargado en la sesión esté guardado
        If typeNameDoc = "ProductDocument" OrElse typeNameDoc = "DrawingDocument" Then
            For Each doc As INFITF.Document In app.Documents
                If Not doc.Saved OrElse String.IsNullOrEmpty(doc.Path) Then
                    If typeNameDoc = "ProductDocument" Then Return CatiaSessionStatus.ProductDocumentNotSaved
                    Return CatiaSessionStatus.DrawingDocumentNotSaved
                End If
            Next

            ' Si pasó el bucle, todo está guardado
            Return If(typeNameDoc = "ProductDocument", CatiaSessionStatus.ProductDocument, CatiaSessionStatus.DrawingDocument)
        End If

        ' 3. Otros tipos de documentos (sin validación profunda)
        Select Case typeNameDoc
            Case "PartDocument" : Return CatiaSessionStatus.PartDocument
            Case "CatalogDocument" : Return CatiaSessionStatus.CatalogDocument
            Case "AnalysisDocument" : Return CatiaSessionStatus.AnalysisDocument
            Case "CATProcessDocument" : Return CatiaSessionStatus.ProcessDocument
            Case Else : Return CatiaSessionStatus.Unknown
        End Select
    End Function



    Public ReadOnly Property Application As INFITF.Application
        Get
            Return _app
        End Get
    End Property


    Public ReadOnly Property Status As CatiaSessionStatus
        Get
            Return _status
        End Get
    End Property


    Public ReadOnly Property IsReady As Boolean
        Get
            Return Status = CatiaSessionStatus.ProductDocument OrElse Status = CatiaSessionStatus.DrawingDocument
        End Get
    End Property

    Public ReadOnly Property Description As String
        Get
            Return "CatiaSessionStatus." & Me.Status.ToString()
        End Get
    End Property

    'Public ReadOnly Property Description As String
    '    Get
    '        Select Case Status
    '            Case CatiaSessionStatus.NotRunning : Return "CATIA is not running."
    '            Case CatiaSessionStatus.NoWindowsOpen : Return "CATIA has no document open."
    '            Case CatiaSessionStatus.ProductDocument : Return "Product document is active and saved."
    '            Case CatiaSessionStatus.ProductDocumentNotSaved : Return "Product document is active but not saved (or new)."
    '            Case CatiaSessionStatus.PartDocument : Return "Part document is active."
    '            Case CatiaSessionStatus.DrawingDocument : Return "Drawing document is active and saved."
    '            Case CatiaSessionStatus.DrawingDocumentNotSaved : Return "Drawing document is active but not saved (or new)."
    '            Case CatiaSessionStatus.CatalogDocument : Return "Catalog document is active."
    '            Case CatiaSessionStatus.AnalysisDocument : Return "Analysis document is active."
    '            Case CatiaSessionStatus.ProcessDocument : Return "Process document is active."
    '            Case Else : Return "Unknown or invalid CATIA state."
    '        End Select
    '    End Get
    'End Property


    Public ReadOnly Property ActiveProductDocument As ProductStructureTypeLib.ProductDocument
        Get
            If Me.Status = CatiaSessionStatus.ProductDocument Then
                Return CType(_app.ActiveDocument, ProductStructureTypeLib.ProductDocument)
            End If
            Return Nothing
        End Get
    End Property


    Public ReadOnly Property ActiveDrawingDocument As DRAFTINGITF.DrawingDocument
        Get
            If _app.ActiveDocument Is Nothing Then Return Nothing
            If TypeName(_app.ActiveDocument) = "DrawingDocument" Then
                Return CType(_app.ActiveDocument, DRAFTINGITF.DrawingDocument)
            End If
            Return Nothing
        End Get
    End Property


    'Public ReadOnly Property ActiveDrawingDocument As DRAFTINGITF.DrawingDocument
    '    Get
    '        If Me.Status = CatiaSessionStatus.DrawingDocument Then
    '            Return CType(_app.ActiveDocument, DRAFTINGITF.DrawingDocument)
    '        End If
    '        Return Nothing
    '    End Get
    'End Property


    Public ReadOnly Property RootProduct As ProductStructureTypeLib.Product
        Get
            Dim doc = Me.ActiveProductDocument
            If doc IsNot Nothing Then
                Return doc.Product
            End If
            Return Nothing
        End Get
    End Property


    Public Enum CatiaSessionStatus
        NotRunning = 0
        NoWindowsOpen = 1
        ProductDocument = 2
        ProductDocumentNotSaved = 9
        PartDocument = 3
        DrawingDocument = 4
        CatalogDocument = 5
        AnalysisDocument = 6
        ProcessDocument = 7
        ScriptDocument = 8
        Unknown = -1
        DrawingDocumentNotSaved = 10
    End Enum

End Class










'Private Function EvaluateStatus(app As INFITF.Application) As CatiaSessionStatus
'    If app Is Nothing Then Return CatiaSessionStatus.NotRunning
'    If app.Windows.Count = 0 Then Return CatiaSessionStatus.NoWindowsOpen

'    Dim oActiveDoc As INFITF.Document = app.ActiveDocument
'    Dim typeNameDoc As String = TypeName(oActiveDoc)

'    If typeNameDoc = "ProductDocument" Then
'        ' 1. Verificamos el Root activo
'        If String.IsNullOrEmpty(oActiveDoc.Path) OrElse Not oActiveDoc.Saved Then
'            Return CatiaSessionStatus.ProductDocumentNotSaved
'        End If

'        ' 2. Verificamos todos los documentos cargados (incluyendo hijos)
'        Try
'            Dim oDocs As INFITF.Documents = app.Documents

'            ' El compilador hará el CType automáticamente por cada elemento
'            For Each subDoc As INFITF.Document In oDocs
'                If Not subDoc.Saved OrElse String.IsNullOrEmpty(subDoc.Path) Then
'                    Return CatiaSessionStatus.ProductDocumentNotSaved
'                End If
'            Next
'        Catch
'            ' En caso de error al iterar (ej. documentos de sistema o procesos en cierre)
'            Return CatiaSessionStatus.Unknown
'        End Try

'        Return CatiaSessionStatus.ProductDocument
'    End If

'    ' Resto de tipos de documentos
'    Select Case typeNameDoc
'        Case "PartDocument" : Return CatiaSessionStatus.PartDocument
'        Case "DrawingDocument" : Return CatiaSessionStatus.DrawingDocument
'        Case "CatalogDocument" : Return CatiaSessionStatus.CatalogDocument
'        Case "AnalysisDocument" : Return CatiaSessionStatus.AnalysisDocument
'        Case "CATProcessDocument" : Return CatiaSessionStatus.ProcessDocument
'        Case "CATScriptDocument" : Return CatiaSessionStatus.ScriptDocument
'        Case Else : Return CatiaSessionStatus.Unknown
'    End Select
'End Function

