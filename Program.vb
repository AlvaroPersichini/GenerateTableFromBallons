Imports System.Collections.Specialized.BitVector32
Imports GenerateTableFromBallons.CatiaSession

Module Program


    Sub Main()


        ' ********************* GenerateTableFromBalloons2 ******************************
        ' Salida: Crea una nueva hoja con una tabla o solo una tabla con el contenido de
        ' los balloons y/o textos que hay en una vista, hoja o documento.
        ' Entrada:
        '           1) oView / oSheet / oDrwDoc
        '           2) strDir: directorio de archivos para propiedades de balloons/textos.
        '           3) getBallons: Booleano para decidir si se toman los balloons
        '           4) getTexts: Booleano para decidir si se toman los textos
        '
        ' Para obtener las propiedades se utiliza la funcion: "ObtenerPropiedades"
        ' Para armar el diccionario de balloons y/o textos se utiliza la funcion: "GetBalloonsData"
        ' Este sub proceso "GenerateTableFromBalloons2", puede tomar los textos que son balloons, 
        ' tambien textos que no son balloons. Se decide con "getBallons" y "getTexts".



        Dim oCatiaSession As New CatiaSession()
        If Not oCatiaSession.IsReady OrElse oCatiaSession.Status <> CatiaSessionStatus.DrawingDocument Then
            Console.WriteLine("Por favor, asegúrese de tener un CATDrawing activo y guardado.")
            Return
        End If



        Dim oDrawingDoc As DRAFTINGITF.DrawingDocument = oCatiaSession.ActiveDrawingDocument
        Dim oDrawingRoot As DRAFTINGITF.DrawingRoot = oDrawingDoc.DrawingRoot
        Dim oSheet As DRAFTINGITF.DrawingSheet = oDrawingRoot.ActiveSheet
        Dim oViews As DRAFTINGITF.DrawingViews = oSheet.Views
        Dim strDir As String = "D:\OneDrive\_CATIA\_V5R21-DLN\FS-1000_R03"
        Dim getBallons As Boolean = True
        Dim getTexts As Boolean = True
        DrawingTools.GenerateTableFromBalloons2(oDrwDoc, strDir, getBallons, getTexts)


    End Sub




End Module
