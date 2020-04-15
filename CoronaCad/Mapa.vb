Imports System.Globalization
Imports System.IO
Imports System.Reflection
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

            ' adicionar leitura direta do servidor
            'datatable com os dados ,

            Dim dt As systemData.DataTable = LerCSV(caminho, ";")

            'o nome das colunas foram alterados ao longo das publicações
            'padrão em portugues
            'regiao  Estado	data	casosNovos	casosAcumulados	obitosNovos	obitosAcumulados

            dt.Columns(0).ColumnName = "regiao"
            dt.Columns(1).ColumnName = "Estado"
            dt.Columns(2).ColumnName = "data"
            dt.Columns(3).ColumnName = "casosNovos"
            dt.Columns(4).ColumnName = "casosAcumulados"
            dt.Columns(5).ColumnName = "obitosNovos"
            dt.Columns(6).ColumnName = "obitosAcumulados"

            Dim nomeEstados() As String = {"AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO"}

            Dim estadosCol As New Collection
            For Each nomeEstado As String In nomeEstados
                estadosCol.Add(New Estado(nomeEstado), nomeEstado)
            Next

            'criar polilinhas dos estados
            Dim polyColl As New Collection
            For Each estado As Estado In estadosCol
                Dim polyEstado As New Polyline With {.Layer = estado.Nome}
                polyEstado.AddVertexAt(polyEstado.NumberOfVertices, New Point2d(0, 0), 0, 0, 0)
                polyColl.Add(polyEstado, estado.Nome)

                model.AppendEntity(polyEstado)
                tr.AddNewlyCreatedDBObject(polyEstado, True)

            Next

            Dim diaInicial As Integer = dt.Compute("Min(data)", "") '43860 primeiro dia do banco de dados
            Dim diaFinal As Integer = dt.Compute("Max(data)", "") '43927 ultimo dia do banco de dados
            Dim distorcaoX As Double = 220 ' escala do grafico

            'texto com a data atual
            Dim textoData As New DBText With
                {.Height = 500,
                .Position = New Point3d(7252, 11000, 1)}
            textoData.SetDatabaseDefaults(db)

            model.AppendEntity(textoData)
            tr.AddNewlyCreatedDBObject(textoData, True)

            'adiciona novo ponto
            For i = diaInicial To diaFinal Step 1 'frequencia de dias

                'mudando data do texto conforme data atual
                'conveter de dias corridos para dd/MM/aa

                Dim dataAtual As DateTime = DateTime.FromOADate(i)
                textoData.TextString = dataAtual.Date

                For Each estado As String In nomeEstados
                    Dim polyEstado As Polyline = polyColl(estado)
                    polyEstado.AddVertexAt(polyEstado.NumberOfVertices, New Point2d((i - diaInicial) * distorcaoX, estadosCol(estado).LocalizaValor("casosAcumulados", polyEstado.Layer, i, dt)), 0, 0, 0)
                Next
                PausaAtualiza(100)
            Next

            tr.Commit()
        End Using

    End Sub
    Private Class Estado

        'estado	data	casosNovos	casosAcumulados	obitosNovos	obitosAcumulados
        Dim _nome As String
        'Dim _data As Date
        'Dim _casosNovos As Integer
        'Dim _casosAcumulados As Integer
        'Dim _obitosNovos As Integer
        'Dim _obitosAcumulados As Integer
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

        Public Function LocalizaValor(nomeColuna As String, estado As String, data As String, dt As systemData.DataTable)

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

            Return Nothing
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
    Private Sub CriarGrafico(dt As systemData.DataTable, estadosCol As Collection, colunaDado As String, frequenciaDias As Integer, tipoDistorcao As String)

        Dim doc As Document = acApp.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed = doc.Editor

        Using tr = doc.TransactionManager.StartTransaction
            'tabela de bloco em escrita
            Dim bt As BlockTable = db.BlockTableId.GetObject(OpenMode.ForWrite)
            'e modelspace em escrita
            Dim model As BlockTableRecord = bt(BlockTableRecord.ModelSpace).GetObject(OpenMode.ForWrite)

            'cria e adiciona a polilinha
            For Each estado As Estado In estadosCol

                'criar coleção de linhas para depois ir adicionando pontos 

            Next

            'Dim linha As New Polyline With {.Layer = Estado.Nome}
            '    model.AppendEntity(linha)
            '    tr.AddNewlyCreatedDBObject(linha, True)

            '    Dim diaInicial As Integer = dt.Compute("Min(data)", "") '43860
            '    Dim diaFinal As Integer = dt.Compute("Max(data)", "") '43927
            '    Dim distorcaoX As Double = 2

            '    linha.AddVertexAt(linha.NumberOfVertices, New Point2d(0, 0), 0, 0, 0)

            '    'adiciona novo ponto
            '    For i = diaInicial To diaFinal Step frequenciaDias
            '        linha.AddVertexAt(linha.NumberOfVertices, New Point2d((i - diaInicial) * distorcaoX, Estado.LocalizaValor(colunaDado, Estado.Nome, i, dt) / 50), 0, 0, 0)

            '    Next
            '    linha.AddVertexAt(linha.NumberOfVertices, New Point2d((diaFinal - diaInicial) * distorcaoX, Estado.LocalizaValor(colunaDado, Estado.Nome, diaFinal, dt) / 50), 0, 0, 0)

            '    tr.Commit()
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
                'algumas publicações do arquivo csv estao com data no formato dd/MM/aaaa e outras no formato dd-MM-aaaa, devera ser convertido para inteiro

                If rows(2).Contains("/") Then
                    Dim convertData As Date = rows(2)
                    rows(2) = convertData.ToOADate
                ElseIf rows(2).Contains("-") Then
                    rows(2).Replace("-", "/")
                    Dim convertData As Date = rows(2)
                    rows(2) = convertData.ToOADate

                End If
                dr(i) = rows(i)
            Next

            dt.Rows.Add(dr)
        End While
        Return dt

    End Function
End Class
