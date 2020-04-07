Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports acApp = Autodesk.AutoCAD.ApplicationServices.Application

Imports systemData = System.Data ' conflito com datatable do cad
Public Class Mapa
    <CommandMethod("CoronaCad")>
    Public Sub CoronaCad()

        Dim doc As Document = acApp.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed = doc.Editor

        Using tr = doc.TransactionManager.StartTransaction
            'tabela de bloco em escrita
            Dim bt As BlockTable = db.BlockTableId.GetObject(OpenMode.ForWrite)
            'e modelspace em escrita
            Dim model As BlockTableRecord = bt(BlockTableRecord.ModelSpace).GetObject(OpenMode.ForWrite)


            Dim dt As systemData.DataTable = LerCSV("C:\Users\raina\Desktop\1f2e9efc2bdd487d4f3b693467aeb925_Download_COVID19_20200406.csv")

            'coleção para os objectos do layer
            Dim ents As New ObjectIdCollection

            'pede o nome do layer
            Dim pr As PromptResult = ed.GetString(vbLf & "Nome do layer: ")

            If pr.Status = PromptStatus.OK Then
                ents = ObjetosLayer(pr.StringResult)
            Else
                Exit Sub
            End If

            For Each estado As ObjectId In ents

                Dim objeto As DBObject = tr.GetObject(estado, OpenMode.ForWrite)

                If objeto.GetRXClass.DxfName = "LWPOLYLINE" Then

                    ''para criar as hachuras do estados
                    'Dim hatchE As New Hatch
                    'Dim dbObjids As New ObjectIdCollection()
                    'dbObjids.Add(estado)
                    'model.AppendEntity(hatchE)
                    'tr.AddNewlyCreatedDBObject(hatchE, True)
                    'With hatchE
                    '    .SetHatchPattern(HatchPatternType.PreDefined, "SOLID")
                    '    .Associative = True
                    '    .AppendLoop(HatchLoopTypes.External, dbObjids)
                    '    .EvaluateHatch(True)
                    '    .Layer = pr.StringResult
                    'End With



                    tr.Commit()
                End If

            Next
        End Using

    End Sub
    ''' <summary>
    ''' Função que retorna todos os objectId do layer
    ''' </summary>
    Private Function ObjetosLayer(nomeLayer As String) As ObjectIdCollection
        Dim doc = acApp.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor

        Using tr As Transaction = doc.TransactionManager.StartOpenCloseTransaction

            'selecionar todos objetos do layer
            Dim tvl As TypedValue() = New TypedValue(0) {New TypedValue(CInt(DxfCode.LayerName), nomeLayer)}
            Dim sf As SelectionFilter = New SelectionFilter(tvl)
            Dim psr As PromptSelectionResult = ed.SelectAll(sf)

            If psr.Status = PromptStatus.OK Then
                'returna a coleção com os objecids dos layer
                Return New ObjectIdCollection((psr.Value.GetObjectIds()))
            Else
                'retorna colecao vazia
                Return New ObjectIdCollection()
            End If
        End Using

    End Function
    Private Sub MudarCorObjeto(corNome As String, tipoObjeto As Type)

    End Sub

    Function LerCSV(ByVal path As String) As System.Data.DataTable

        Try

            Dim sr As New StreamReader(path)
            Dim fullFileStr As String = sr.ReadToEnd()
            sr.Close()
            sr.Dispose()

            Dim lines As String() = fullFileStr.Split(ControlChars.Lf)
            Dim dt As New systemData.DataTable()
            Dim sArr As String() = lines(0).Split(";"c)


            For Each s As String In sArr
                dt.Columns.Add(New systemData.DataColumn())
            Next
            Dim row As DataRow
            Dim finalLine As String = ""

            For Each line As String In lines
                row = dt.NewRow()
                finalLine = line.Replace(Convert.ToString(ControlChars.Cr), "")
                row.ItemArray = finalLine.Split(";"c)
                dt.Rows.Add(row)
            Next

            MsgBox(sArr.Count)
            Dim count As Integer = 0
            For Each titulo As String In sArr
                MsgBox(titulo)
                dt.Columns(count).ColumnName = titulo
                count += 1

            Next

            For Each colunas As systemData.DataColumn In dt.Columns
                MsgBox("nome col: " & colunas.ColumnName)
            Next

            Return dt
        Catch ex As Exception

            Throw ex
        End Try

    End Function
End Class
