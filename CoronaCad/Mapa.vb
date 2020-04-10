Imports System.IO

Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
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

            'abrir aquivo
            Dim openFileDialog1 As OpenFileDialog = New OpenFileDialog()
            Dim caminho As String
            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                caminho = openFileDialog1.FileName
            Else
                Exit Sub
            End If

            'datatable com os dados
            Dim dt As systemData.DataTable = LerCSV(caminho, ";")

            Dim PR As Estado = New Estado("PR")

            Dim data As String = "43927"

            'CriarGrafico(dt, "PR", dt.Columns(0), dt.Columns(1))
            'tr.Commit()
        End Using

    End Sub
    Private Class Estado

        'estado	data	casosNovos	casosAcumulados	obitosNovos	obitosAcumulados
        Dim _nome As String
        Dim _data As Date
        Dim _casosNovos As Integer
        Dim _casosAcumulados As Integer
        Dim _obitosNovos As Integer
        Dim _obitosAcumulados As Integer
        Public Sub New(ByVal abreviacao As String)
            _nome = abreviacao
        End Sub

        Public Function GetCasosNovos(data As Integer, dt As systemData.DataTable)
            Return LocalizaValor("casosNovos", _nome, data, dt)
        End Function
        Public Function GetCasosAcumulados(data As Integer, dt As systemData.DataTable)
            Return LocalizaValor("casosAcumulados", _nome, data, dt)
        End Function
        Public Function GetObitosNovos(data As Integer, dt As systemData.DataTable)
            Return LocalizaValor("obitosNovos", _nome, data, dt)
        End Function
        Public Function GetObitosAcumulados(data As Integer, dt As systemData.DataTable)
            Return LocalizaValor("obitosAcumulados", _nome, data, dt)
        End Function

        Private Function LocalizaValor(nomeColuna As String, estado As String, data As String, dt As systemData.DataTable)

            ' filtro coluna "data" = a data informada
            Dim filtro As String = "data" & "=" & data

            ' localiza as linhas com o parametro do  filtro
            Dim drc As DataRow() = dt.Select(filtro)

            'index da coluna procurada
            Dim indexColunaProcurada As Integer = dt.Columns(nomeColuna).Ordinal
            'percorre todoas as linhas encontradas
            For Each linha As systemData.DataRow In drc
                'pegar apenas o estado selecionado
                If linha.Item(1).ToString = estado Then
                    Return linha(indexColunaProcurada)
                End If
            Next

        End Function
        Property Nome As String
            Get
                Return _nome
            End Get

            Set(ByVal Value As String)
                _nome = Value
            End Set
        End Property

    End Class

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
    Private Sub CriarGrafico(dt As systemData.DataTable, estado As String, colunaX As systemData.DataColumn, colunaY As systemData.DataColumn)

        Dim doc As Document = acApp.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed = doc.Editor

        Using tr = doc.TransactionManager.StartTransaction
            'tabela de bloco em escrita
            Dim bt As BlockTable = db.BlockTableId.GetObject(OpenMode.ForWrite)
            'e modelspace em escrita
            Dim model As BlockTableRecord = bt(BlockTableRecord.ModelSpace).GetObject(OpenMode.ForWrite)

            'cria e adiciona a polilinha
            Dim linha As New Polyline With {.Layer = estado}
            model.AppendEntity(linha)
            tr.AddNewlyCreatedDBObject(linha, True)

            'adiciona novo ponto
            For Each dtLinha As DataRow In dt.Rows

                linha.AddVertexAt(linha.NumberOfVertices, New Point2d(dtLinha.ItemArray(0), dtLinha.ItemArray(1)), 0, 0, 0)
                PausaAtualiza(500)
            Next
            tr.Commit()
        End Using

    End Sub
    Private Sub MudarCorObjeto(corNome As String, tipoObjeto As Type)

    End Sub
    Private Sub PausaAtualiza(tempo As Integer)
        Dim doc As Document = acApp.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed = doc.Editor

        doc.TransactionManager.QueueForGraphicsFlush()
        doc.TransactionManager.FlushGraphics()
        ed.UpdateScreen()
        System.Threading.Thread.Sleep(tempo)
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
