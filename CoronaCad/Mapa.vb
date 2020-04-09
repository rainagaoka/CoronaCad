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

            Dim dt As systemData.DataTable = LerCSV("C:\Users\raina\Desktop\1f2e9efc2bdd487d4f3b693467aeb925_Download_COVID19_20200406.csv", ";")

            MsgBox("L: " & dt.Rows.Count & "C: " & dt.Columns.Count)

            For Each item As systemData.DataColumn In dt.Columns
                MsgBox(item.ColumnName)

            Next

            For Each item As String In dt.Rows(1).ItemArray
                MsgBox(item)

            Next


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

    ''' <summary>
    ''' Função que converte um arquivo de texto com delimitador, a primeira linha deve conter o título das colunas.
    ''' </summary>
    Public Function LerCSV(ByVal strFilePath As String, delimitador As String) As System.Data.DataTable

        'abre o arquivo
        Dim sr As StreamReader = New StreamReader(strFilePath)
        'pega a primeira linha como cabeçalho

        Dim headers As String() = sr.ReadLine().Split(delimitador)
        Dim dt As systemData.DataTable = New systemData.DataTable()

        'cria as colunas conforme primeira linha 
        For Each header As String In headers
            dt.Columns.Add(header)
        Next

        'le o restante até o final
        While Not sr.EndOfStream
            Dim rows As String() = sr.ReadLine().Split(delimitador)
            Dim dr As DataRow = dt.NewRow()

            For i As Integer = 0 To headers.Length - 1
                dr(i) = rows(i)
            Next

            dt.Rows.Add(dr)
        End While

        Return dt

    End Function
End Class
