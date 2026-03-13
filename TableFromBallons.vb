Module TableFromBallons

    Public Structure PropertiesContainer

        Public sPN As String
        Public sDescription As String
        Public sDefinition As String
        Public sSource As String

    End Structure

    Public Function ObtenerPropiedades(oAppCATIA As INFITF.Application, strDir As String) As Dictionary(Of String, PropertiesContainer)
        Dim oFileSystem As INFITF.FileSystem = oAppCATIA.FileSystem
        Dim oFolder As INFITF.Folder = oFileSystem.GetFolder(strDir)
        Dim iDocument As INFITF.Document
        Dim iPartDoc As MECMOD.PartDocument
        Dim iProductDoc As ProductStructureTypeLib.ProductDocument
        Dim oDic As New Dictionary(Of String, PropertiesContainer)

        For Each file As INFITF.File In oFolder.Files
            If file.Type = "CATIA Part" Then

                ' Se lee el archivo, se convierte a document y se convierte a PartDocument
                iDocument = oAppCATIA.Documents.Read(file.Path)
                iPartDoc = CType(iDocument, MECMOD.PartDocument)

                ' Le cargo todo al struct
                Dim oPropertiesContainer As New PropertiesContainer With {
                    .sPN = iPartDoc.Product.ReferenceProduct.PartNumber,
                    .sDescription = iPartDoc.Product.ReferenceProduct.DescriptionRef,
                    .sDefinition = iPartDoc.Product.ReferenceProduct.Definition,
                    .sSource = iPartDoc.Product.ReferenceProduct.Source
                }

                ' Si no existe ya el key entonces le cargo al oDic el PN en el Key y el struct en el Value
                If Not oDic.ContainsKey(iPartDoc.Product.ReferenceProduct.PartNumber) Then
                    oDic.Add(iPartDoc.Product.ReferenceProduct.PartNumber, oPropertiesContainer)
                End If

            ElseIf file.Type = "CATIA Product" Then

                iDocument = oAppCATIA.Documents.Read(file.Path)
                iProductDoc = CType(iDocument, ProductStructureTypeLib.ProductDocument)

                ' Le cargo todo al struct
                Dim oPropertiesContainer As New PropertiesContainer With {
                    .sPN = iProductDoc.Product.ReferenceProduct.PartNumber,
                    .sDescription = iProductDoc.Product.ReferenceProduct.DescriptionRef,
                    .sDefinition = iProductDoc.Product.ReferenceProduct.Definition,
                    .sSource = iProductDoc.Product.ReferenceProduct.Source
                }
                ' Si no existe ya el key entonces le cargo al oDic el PN en el Key y el struct en el Value
                If Not oDic.ContainsKey(iProductDoc.Product.ReferenceProduct.PartNumber) Then
                    oDic.Add(iProductDoc.Product.ReferenceProduct.PartNumber, oPropertiesContainer)
                End If
            End If
        Next

        ' --- Bloque para escribir las keys en un archivo txt ---
        Dim outputPath As String = "C:\Temp\PropiedadesKeys.txt"
        Dim sb As New System.Text.StringBuilder()

        sb.AppendLine("Listado de Keys del Diccionario de Propiedades:")
        sb.AppendLine("================================================")
        sb.AppendLine("")

        For Each key As String In oDic.Keys
            sb.AppendLine(key)
        Next

        sb.AppendLine("")
        sb.AppendLine("Total de elementos: " & oDic.Count)

        IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(outputPath))
        IO.File.WriteAllText(outputPath, sb.ToString())
        ' --- Fin del bloque ---

        Return oDic


    End Function


    ' ********************* GetBalloonsData  *****************
    ' Objetivo: arma un oDic con la primera linea de texto de los
    ' balloons o textos insertados con flecha que existen en una vista,
    ' una sheet o en todo un documento de tipo CATDrawing.
    ' Este diccionario de balloons luego se lo usa como argumento para la funcion que genera una tabla.
    ' Las variables getBalloons y getTexs son para indicar si se quieren obtener los balloons, los textos o ambos.
    Public Function GetBalloonsData(oDrwDocument As DRAFTINGITF.DrawingDocument, getBallons As Boolean, getText As Boolean) As System.Collections.Specialized.StringDictionary


        Dim oSheets As DRAFTINGITF.DrawingSheets = oDrwDocument.Sheets
        Dim oViews As DRAFTINGITF.DrawingViews
        Dim oDicBalloons As New System.Collections.Specialized.StringDictionary
        Dim strReader As IO.StringReader
        Dim aLine As String ' es la primer línea del contenido de texto de los balloons


        For Each sheet As DRAFTINGITF.DrawingSheet In oSheets

            oViews = sheet.Views

            For Each view As DRAFTINGITF.DrawingView In oViews
                ' evitar "Main View" y  "Background View" 
                If view.Name = "Main View" Or view.Name = "Background View" Then
                    Continue For
                End If

                ' evita las vistas que no tienen DrawingText
                If view.Texts.Count = 0 Then
                    Continue For
                End If

                For Each iDrawingText As DRAFTINGITF.DrawingText In view.Texts

                    ' Si no se quieren obtener los balloons
                    If Left(iDrawingText.Name, 7) = "Balloon" And getBallons = False Then
                        Continue For
                    End If

                    ' Si no se quieren obtener los textos
                    If Left(iDrawingText.Name, 7) <> "Balloon" And getText = False Then
                        Continue For
                    End If

                    ' Evita textos vacios
                    If iDrawingText.Text = "" Then
                        Continue For
                    End If


                    ' Obtengo solo la primer linea del balloon, la cual contiene el PartNumber
                    strReader = New IO.StringReader(iDrawingText.Text)
                    aLine = strReader.ReadLine


                    ' Saltea los que ya ha computado
                    If oDicBalloons.ContainsKey(aLine) Then
                        Continue For
                    End If

                    oDicBalloons.Add(aLine, iDrawingText.Name)

                Next
            Next
        Next

        Return oDicBalloons

    End Function


    ' ********************* GetBalloons Data  *****************
    ' Objetivo: arma un oDic con la primera linea de texto de los balloons que existen en una vista,
    ' una sheet o en todo un documento de tipo CATDrawing.
    ' Este diccionario de balloons luego se lo usa como argumento para la funcion que genera 
    ' una tabla.
    Public Function GetBalloonsData(oSheet As DRAFTINGITF.DrawingSheet) As System.Collections.Specialized.StringDictionary

        Dim oAppCATIA As INFITF.Application = oSheet.Application
        Dim oDocuments As INFITF.Documents = oAppCATIA.Documents
        Dim oActiveDocument As INFITF.Document = oAppCATIA.ActiveDocument
        Dim oViews As DRAFTINGITF.DrawingViews = oSheet.Views
        Dim oDicBalloons As New System.Collections.Specialized.StringDictionary
        Dim strReader As IO.StringReader
        Dim aLine As String ' es la primer línea del contenido de texto de los balloons

        ' ****************************************************************************************************
        ' Todo este bloque de codigo es para calcular la cantidad de filas que tiene que tener la tabla
        ' La cantindad de filas es igual a la cantidad de balloons, para encontrar la cantidad de balloons
        ' el proceso que se sigue es el siguiente:
        ' 1) Recorrido de todas las vistas de una determinada oSheet, evitando "Main View" y  "Background View" 
        ' 2) Evita las vistas que no tienen DrawingText
        ' 3) Solo tener en cuenta los drawingText que sean "balloon"
        ' 4) Evitar los textos que esten en blanco
        ' IMPORTANTE: solo se van a tener en cuenta los balloons. Si se ha insertado texto de forma manual,
        ' no va a ser tenido en cuenta.

        For Each view As DRAFTINGITF.DrawingView In oViews
            ' evitar "Main View" y  "Background View" 
            If view.Name = "Main View" Or view.Name = "Background View" Then
                Continue For
            End If
            If view.Texts.Count = 0 Then ' evita las vistas que no tienen DrawingText
                Continue For
            End If
            For Each iDrawingText As DRAFTINGITF.DrawingText In view.Texts


                ' Al agregar un balloon, ese nuevo drawingText generado tiene como nombre "Balloon.x", al contrario 
                ' de agregar un texto, que en este caso el DrawingText generado tiene en la prop name = "Text.x"
                ' Entonces se puede usar esto para computar uno u otro, o ambos.

                'Si se quiere computar tanto text como balloons entonces dejar todo comentado.

                'Si no se quiere computar los textos que son balloons hay que descomentar esto
                'If Left(iDrawingText.Name, 7) = "Balloon" Then ' Solo los textos que son balloons
                '    Continue For
                'End If

                ' Si no se quiere computar los DrawingText que tienen como nombre "text" hay que descomentar esto
                'If Left(iDrawingText.Name, 4) = "Text" Then ' Solo los textos que son balloons
                '    Continue For
                'End If


                If iDrawingText.Text = "" Then ' Solo los textos que no estan vacios
                    Continue For
                End If


                ' Obtengo solo la primer linea del Balloon/Text, la cual contiene el PartNumber
                strReader = New IO.StringReader(iDrawingText.Text)
                aLine = strReader.ReadLine


                ' Saltea los que ya ha computado
                If oDicBalloons.ContainsKey(aLine) Then
                    Continue For
                End If


                ' Agrega la primer linea del text (que deberia ser el PartNumber)
                oDicBalloons.Add(aLine, iDrawingText.Name)

            Next
        Next

        Return oDicBalloons

    End Function

    Sub GenerateTableFromBalloons2(oDrwDoc As DRAFTINGITF.DrawingDocument, strDis As String, getBallons As Boolean, getText As Boolean)

        ' Referencia las variables de interfaz
        Dim oAppCATIA As INFITF.Application = oDrwDoc.Application
        oAppCATIA.DisplayFileAlerts = False
        Dim oSheets As DRAFTINGITF.DrawingSheets = oDrwDoc.Sheets


        ' Arma el diccionario con los archivos contenidos en el directorio que se le indique
        ' computa todos los archivos. Esto si bien funciona, esta computando todo, y debería computar solo los que están en el drawing.
        Dim oDicProperties As Dictionary(Of String, PropertiesContainer) = ObtenerPropiedades(oAppCATIA, strDis)


        ' Arma el diccionario del contenido de los balloons
        Dim oDicBalloons As Specialized.StringDictionary = GetBalloonsData(oDrwDoc, getBallons, getText)


        ' Agrego una hoja de lista de Piezas (despues de haber armado el diccionario)
        Dim oSheetBOM As DRAFTINGITF.DrawingSheet = oSheets.Add("Lista de Piezas")


        ' Vista de la tabla de materiales
        Dim oViews As DRAFTINGITF.DrawingViews = oSheetBOM.Views
        Dim oViewListaMateriles As DRAFTINGITF.DrawingView = oViews.Add("Tabla")
        oViewListaMateriles.Activate()


        ' Creacion de la tabla
        Dim oTables As DRAFTINGITF.DrawingTables = oViewListaMateriles.Tables
        Dim oTable As DRAFTINGITF.DrawingTable = oTables.Add(0, 0, oDicBalloons.Count + 1, 3, 5, 100)


        '******************************** Insercion de los datos *******************************************

        Dim rawNumber As Integer = 2 ' empieza desde el segundo renglon, ya que el primero son las cabeceras
        Dim strPN As String

        ' Insercion de los textos de cabecera
        oTable.SetCellString(1, 1, "N° de Parte")
        oTable.SetCellString(1, 2, "Descripción")
        oTable.SetCellString(1, 3, "Código")


        ' IMPORTANTE:
        ' las key estan todas en mniniscula, hay que pasarlas a mayuscula para que coincidan con las del diccionario de propiedades
        ' Esto es peligroso, habría que estandarizar el tema de mayusculas y minusculas en los dos diccionarios

        ' insercion de los no NCU en la tabla
        For Each de As DictionaryEntry In oDicBalloons
            If Not oDicProperties.ContainsKey(UCase(de.Key.ToString)) Then
                Continue For
            End If
            strPN = UCase(de.Key.ToString)
            If strPN.Substring(0, 3) = "NCU" Then
                Continue For
            End If
            oTable.SetCellString(rawNumber, 1, UCase(de.Key.ToString))
            oTable.SetCellString(rawNumber, 2, oDicProperties.Item(UCase(de.Key.ToString)).sDescription)
            oTable.SetCellString(rawNumber, 3, oDicProperties.Item(UCase(de.Key.ToString)).sDefinition)
            rawNumber += 1
        Next

        ' insercion de los NCU en la tabla
        For Each de As DictionaryEntry In oDicBalloons
            If Not oDicProperties.ContainsKey(UCase(de.Key.ToString)) Then
                Continue For
            End If
            strPN = UCase(de.Key.ToString)
            If strPN.Substring(0, 3) <> "NCU" Then
                Continue For
            End If
            oTable.SetCellString(rawNumber, 1, UCase(de.Key.ToString))
            oTable.SetCellString(rawNumber, 2, oDicProperties.Item(UCase(de.Key.ToString)).sDescription)
            oTable.SetCellString(rawNumber, 3, oDicProperties.Item(UCase(de.Key.ToString)).sDefinition)
            rawNumber += 1
        Next


        ' Formato de la tabla
        TableFormat(oTable)


        ' Esto es para poner el el vertice izq inferior en el (0,0) de la vista de la tabla.
        Dim acumulador As Single = 0
        For i = 1 To oDicBalloons.Count + 1
            acumulador += oTable.GetRowSize(i)
        Next
        oTable.y = acumulador


        oSheetBOM.Update()

    End Sub


    Private Sub TableFormat(oTable As DRAFTINGITF.DrawingTable)

        Dim textoInterno As DRAFTINGITF.DrawingText
        Dim oDrawingTextProperties As DRAFTINGITF.DrawingTextProperties

        oTable.ComputeMode = DRAFTINGITF.CatTableComputeMode.CatTableComputeOFF

        ' Nombre
        oTable.Name = "Tabla de Piezas"


        ' Textos:
        For colNum As Integer = 1 To oTable.NumberOfColumns
            For rawNum As Integer = 1 To oTable.NumberOfRows ' empieza del indice 2 porque en la cabecera va otro formato
                textoInterno = oTable.GetCellObject(rawNum, colNum)
                oDrawingTextProperties = textoInterno.TextProperties
                oDrawingTextProperties.FontName = "Arial (TrueType)"
                oDrawingTextProperties.FontSize = 3.5
                oDrawingTextProperties.Justification = DRAFTINGITF.CatJustification.catCenter
            Next
        Next



        ' ********************************  Bordes ****************************************
        ' Limpiar bordes primero
        For colNum As Integer = 1 To oTable.NumberOfColumns
            For rawNum As Integer = 1 To oTable.NumberOfRows
                oTable.SetCellBorderType(rawNum, colNum, DRAFTINGITF.CatTableBorderType.CatTableNone)
            Next
        Next

        ' Seatea los bordes
        For colNum As Integer = 1 To oTable.NumberOfColumns
            For rawNum As Integer = 3 To oTable.NumberOfRows
                oTable.SetCellBorderType(rawNum, colNum, DRAFTINGITF.CatTableBorderType.CatTableInside)
            Next
        Next



        ' Textos cabecera:
        Dim oCabecera1 As DRAFTINGITF.DrawingText
        For i = 1 To 3
            oCabecera1 = oTable.GetCellObject(1, i)
            oCabecera1.SetFontName(0, 0, "Arial (TrueType)")
            oDrawingTextProperties = oCabecera1.TextProperties
            oDrawingTextProperties.FontSize = 6
            oDrawingTextProperties.Italic = True
            oDrawingTextProperties.Underline = True
            oDrawingTextProperties.Justification = DRAFTINGITF.CatJustification.catCenter
        Next


        ' Tamaño
        oTable.SetColumnSize(1, 60)
        oTable.SetColumnSize(2, 77)
        oTable.SetColumnSize(3, 67)


        ' Activa la tabla
        oTable.ComputeMode = DRAFTINGITF.CatTableComputeMode.CatTableComputeON


    End Sub


    Sub GenerateTableFromBalloons2(oSheet As DRAFTINGITF.DrawingSheet, strDis As String)

        ' Referencia las variables de interfaz
        Dim oAppCATIA As INFITF.Application = oSheet.Application
        Dim oDocuments As INFITF.Documents = oAppCATIA.Documents
        Dim oActiveDocument As INFITF.Document = oAppCATIA.ActiveDocument


        ' Arma el diccionario con los archivos contenidos en el directorio que se le indique
        Dim oDicProperties As Dictionary(Of String, PropertiesContainer) = ObtenerPropiedades(oAppCATIA, strDis)


        ' Arma el diccionario del contenido de los balloons
        Dim oDicBalloons As System.Collections.Specialized.StringDictionary = GetBalloonsData(oSheet)



        ' Vista de la tabla de materiales
        Dim oViews As DRAFTINGITF.DrawingViews = oSheet.Views
        Dim oViewListaMateriles As DRAFTINGITF.DrawingView = oViews.Add("BOM")
        oViewListaMateriles.Activate()


        ' Creacion de la tabla
        Dim oTables As DRAFTINGITF.DrawingTables = oViewListaMateriles.Tables
        Dim oTable As DRAFTINGITF.DrawingTable = oTables.Add(0, (oDicBalloons.Count + 1) * 5, oDicBalloons.Count + 1, 3, 5, 100)



        '******************************** Insercion de los datos *******************************************
        Dim rawNumber As Integer = 2 ' empieza desde el segundo renglon, ya que el primero son las cabeceras
        Dim strPN As String

        ' Insercion de los textos de cabecera
        oTable.SetCellString(1, 1, "N° de Parte")
        oTable.SetCellString(1, 2, "Descripción")
        oTable.SetCellString(1, 3, "Código")


        ' insercion de los no NCU en la tabla
        For Each de As DictionaryEntry In oDicBalloons
            strPN = UCase(de.Key.ToString)
            If strPN.Substring(0, 3) <> "NCU" Then
                oTable.SetCellString(rawNumber, 1, UCase(de.Key.ToString))
                oTable.SetCellString(rawNumber, 2, oDicProperties.Item(UCase(de.Key.ToString)).sDescription)
                oTable.SetCellString(rawNumber, 3, oDicProperties.Item(UCase(de.Key.ToString)).sDefinition)
                rawNumber += 1
            End If
        Next

        ' insercion de los NCU en la tabla
        For Each de As DictionaryEntry In oDicBalloons
            strPN = UCase(de.Key.ToString)
            If strPN.Substring(0, 3) = "NCU" Then
                oTable.SetCellString(rawNumber, 1, UCase(de.Key.ToString))
                oTable.SetCellString(rawNumber, 2, oDicProperties.Item(UCase(de.Key.ToString)).sDescription)
                oTable.SetCellString(rawNumber, 3, oDicProperties.Item(UCase(de.Key.ToString)).sDefinition)
                rawNumber += 1
            End If
        Next


        ' Formato de la tabla
        TableFormat(oTable)


        ' Esto es para poner el el vertice izq inferior en el (0,0) de la vista de la tabla.
        Dim acumulador As Single = 0
        For i = 1 To oDicBalloons.Count + 1
            acumulador += oTable.GetRowSize(i)
        Next
        oTable.y = acumulador


        oSheet.Update()

    End Sub









End Module
