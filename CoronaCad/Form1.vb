Imports System.Globalization
Imports System.IO
Imports System.Reflection
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports acApp = Autodesk.AutoCAD.ApplicationServices.Application
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.Interop


Imports systemData = System.Data ' conflito com datatable do cad
Public Class Form1
    <CommandMethod("CoronaCad")>
    Public Sub CoronaCad()

        'marcar opções padrão
        RadioButton_CasosAcumulados.Checked = True
        RadioButton_Linear.Checked = True

        'mostra o form
        Me.ShowDialog()

    End Sub
    Public Sub ZoomView(min2d As Point2d, max2d As Point2d)

        Dim doc As Document = acApp.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed = doc.Editor

        Dim view As New ViewTableRecord With {
            .CenterPoint = min2d + ((max2d - min2d) / 2.0),
            .Height = max2d.Y - min2d.Y,
            .Width = max2d.X - min2d.X
        }

        ed.SetCurrentView(view)

    End Sub
    Private Sub Plotar()
        Dim doc As Document = acApp.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed = doc.Editor

        Using tr = doc.TransactionManager.StartTransaction
            'tabela de bloco em escrita
            Dim bt As BlockTable = db.BlockTableId.GetObject(OpenMode.ForWrite)
            'e modelspace em escrita
            Dim model As BlockTableRecord = bt(BlockTableRecord.ModelSpace).GetObject(OpenMode.ForWrite)

            'abrir aquivo
            Dim caminho As String

            caminho = TextBox_Caminho.Text

            ' !!!! adicionar leitura direta do servidor

            'qual dado selecionado será plotado
            Dim dadoSelecionado As String = "casosAcumulados"
            For Each radiobutton As RadioButton In GroupBox_Dados.Controls.OfType(Of RadioButton)
                If radiobutton.Checked Then
                    Select Case radiobutton.Text
                        Case "Casos Novos"
                            dadoSelecionado = "casosNovos"
                        Case "Casos Acumulados"
                            dadoSelecionado = "casosAcumulados"
                        Case "Óbitos Novos"
                            dadoSelecionado = "obitosNovos"
                        Case "Óbitos Acumulados"
                            dadoSelecionado = "obitosAcumulados"
                    End Select
                End If
            Next

            'em qual escala será plotado o grafico
            Dim escalaEixoY As String = "Linear"
            For Each radiobutton As RadioButton In GroupBox_Escala.Controls.OfType(Of RadioButton)
                Select Case radiobutton.Text
                    Case "Linear"
                        escalaEixoY = "Linear"
                    Case "Logarítmica"
                        escalaEixoY = "Logarítmica"
                End Select
            Next

            'datatable com os dados
            Dim dt As systemData.DataTable = LerCSV(caminho, ";")

            'o nome das colunas foram alterados ao longo das publicações
            'padrão em portugues
            'regiao  Estado	data	casosNovos	casosAcumulados	obitosNovos	obitosAcumulados

            dt.Columns(0).ColumnName = "regiao"
            dt.Columns(1).ColumnName = "estado"
            dt.Columns(2).ColumnName = "data"
            dt.Columns(3).ColumnName = "casosNovos"
            dt.Columns(4).ColumnName = "casosAcumulados"
            dt.Columns(5).ColumnName = "obitosNovos"
            dt.Columns(6).ColumnName = "obitosAcumulados"

            Dim nomeEstados() As String = {"AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO"}

            'cria uma coleção com todas os objetos ESTADO e cria o layer com o nomedo estado
            Dim estadosCol As New Collection
            Dim countColor As Integer = 10 'iniciar na cor 10
            For i = 0 To nomeEstados.Count - 1
                Dim estado As New Estado(nomeEstados(i))
                estadosCol.Add(estado, estado.Nome)
                CriarLayer(estado.Nome, "Continuous", countColor)
                countColor += 9 'variação das cores por estado
            Next

            Dim diaInicial As Integer = dt.Compute("Min(data)", "") '43860 primeiro dia do banco de dados
            Dim diaFinal As Integer = dt.Compute("Max(data)", "") '43927 ultimo dia do banco de dados
            Dim distorcaoX As Double = 220 ' escala do eixo x

            DesenharGráfico(dt, dadoSelecionado, escalaEixoY, distorcaoX)

            'texto com a data atual
            Dim textoData As New DBText With
            {.Height = 500,
            .Position = New Point3d(7252, 11000, 0)}
            textoData.SetDatabaseDefaults(db)

            'model.AppendEntity(textoData)
            'tr.AddNewlyCreatedDBObject(textoData, True)

            AddToModelSpace(textoData, db)

            'criar polilinhas dos estados
            Dim polyColl As New Collection
            For Each estado As Estado In estadosCol
                Dim polyEstado As New Polyline With {.Layer = estado.Nome}
                polyEstado.AddVertexAt(polyEstado.NumberOfVertices, New Point2d(0, 0), 0, 0, 0)
                polyColl.Add(polyEstado, estado.Nome)

                model.AppendEntity(polyEstado)
                tr.AddNewlyCreatedDBObject(polyEstado, True)

            Next

            'plotar pontos e legenda
            For dia = diaInicial To diaFinal Step 1 'frequencia de dias

                'mudando data do texto conforme data atual
                'conveter de dias corridos para dd/MM/aa

                Dim dataAtual As DateTime = DateTime.FromOADate(dia)
                textoData.TextString = dataAtual.Date

                For Each estado As String In nomeEstados
                    Dim polyEstado As Polyline = polyColl(estado)
                    polyEstado.AddVertexAt(polyEstado.NumberOfVertices, New Point2d((dia - diaInicial) * distorcaoX, estadosCol(estado).LocalizaValor(dadoSelecionado, polyEstado.Layer, dia, dt)), 0, 0, 0)
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
    Private Sub DesenharGráfico(dt As systemData.DataTable, tipoDeDado As String, escalaEixoY As String, distorcaoEixoX As Double)
        Dim doc As Document = acApp.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed = doc.Editor

        Using tr = doc.TransactionManager.StartTransaction

            'tabela de bloco em escrita
            Dim bt As BlockTable = db.BlockTableId.GetObject(OpenMode.ForWrite)
            'e modelspace em escrita
            Dim model As BlockTableRecord = bt(BlockTableRecord.ModelSpace).GetObject(OpenMode.ForWrite)

            'data inicial e final
            Dim diaInicial As Integer = dt.Compute("Min(data)", "") '43860 primeiro dia do banco de dados
            Dim diaFinal As Integer = dt.Compute("Max(data)", "") '43927 ultimo dia do banco de dados

            'maior valor do dado selecioando
            Dim maiorValorY = Convert.ToInt32(dt.Compute("Max(" & tipoDeDado & ")", String.Empty))
            'valor para manter a altura do eixo em 10000
            Dim distorçãoEixoY As Double = 10000 / maiorValorY

            'desenhar eixo x e y com base nas datas
            'desenhar exo x
            Dim layerEixoXNome As String = "EixoGrafico"
            CriarLayer(layerEixoXNome, "Continuous", 0)
            Dim eixoXPoly As New Polyline With {.Layer = layerEixoXNome}

            'adiciona os pontos 0,0 e datafinal,0
            eixoXPoly.AddVertexAt(eixoXPoly.NumberOfVertices, New Point2d(0, 0), 0, 0, 0)
            eixoXPoly.AddVertexAt(eixoXPoly.NumberOfVertices, New Point2d((diaFinal - diaInicial) * distorcaoEixoX, 0), 0, 0, 0)

            AddToModelSpace(eixoXPoly, db)

            'desenhar eixo y
            Dim layerEixoYNome As String = "Eixo Grafico"
            CriarLayer(layerEixoXNome, "Continuous", 0)
            Dim eixoYPoly As New Polyline With {.Layer = layerEixoXNome}

            'adiciona os pontos 0,0 e datafinal,0
            eixoYPoly.AddVertexAt(eixoYPoly.NumberOfVertices, New Point2d(0, 0), 0, 0, 0)
            eixoYPoly.AddVertexAt(eixoYPoly.NumberOfVertices, New Point2d(0, maiorValorY), 0, 0, 0)

            AddToModelSpace(eixoYPoly, db)

            'dsenhar linha tracejada dos valores de y e texto valores
            Dim layerEixosSecunY As String = "Eixos Secundarios"
            Dim layerValoresY As String = "Valores Y"

            CriarLayer(layerEixosSecunY, "DASHED2", 250)
            CriarLayer(layerValoresY, "Continuous", 0)
            Dim interEixoSecunY As Integer = 1000
            For i = 1000 To maiorValorY Step interEixoSecunY
                Dim eixoSecun As New Polyline With {.Layer = layerEixosSecunY,
                .LinetypeScale = 2000}

                'adicionar as linhas horizontais 
                eixoSecun.AddVertexAt(eixoSecun.NumberOfVertices, New Point2d(0, i), 0, 0, 0)
                eixoSecun.AddVertexAt(eixoSecun.NumberOfVertices, New Point2d((diaFinal - diaInicial) * distorcaoEixoX, i), 0, 0, 0)

                AddToModelSpace(eixoSecun, db)

                'texto valor eixo y
                Dim valoresYTexto As New DBText With
                {.Height = 500,
                .TextString = i,
                .VerticalMode = TextVerticalMode.TextVerticalMid,
                .HorizontalMode = TextHorizontalMode.TextRight,
                .AlignmentPoint = New Point3d(-400, i, 0),
                .Layer = layerValoresY
                }

                'adicona ao database e a transaçção
                AddToModelSpace(valoresYTexto, db)

                'por alguma razao precisa disso para ficar correto na tela durante a execucao
                valoresYTexto.AdjustAlignment(db)

            Next

            'desenhar tick dos valores de x e texto com as datas
            Dim layerValoresX As String = "Valores X"
            CriarLayer(layerValoresX, "Continuous", 0)

            Dim layerTickEixoX As String = "Tick Eixo X"
            CriarLayer(layerTickEixoX, "Continuous", 250)

            'percorre os dias
            For i = diaInicial To diaFinal Step 7

                Dim tickEixoX As New Polyline With {.Layer = layerTickEixoX}

                'adicionar as linhas horizontais 
                tickEixoX.AddVertexAt(tickEixoX.NumberOfVertices, New Point2d((i - diaInicial) * distorcaoEixoX, 0), 0, 0, 0)
                tickEixoX.AddVertexAt(tickEixoX.NumberOfVertices, New Point2d((i - diaInicial) * distorcaoEixoX, -300), 0, 0, 0)

                AddToModelSpace(tickEixoX, db)

                'texto valor eixo y
                Dim dataAtual As DateTime = DateTime.FromOADate(i)
                Dim dataFormatada As String = Format(dataAtual, "dd-MMM") ' data no formatao 01-jan

                Dim valoresXTexto As New DBText With
                {.Height = 500,
                .TextString = dataFormatada,
                .VerticalMode = TextVerticalMode.TextVerticalMid,
                .HorizontalMode = TextHorizontalMode.TextRight,
                .AlignmentPoint = New Point3d((i - diaInicial) * distorcaoEixoX, -500, 0),
                .Rotation = 45
                }

                'adicona ao database e a transaçção
                AddToModelSpace(valoresXTexto, db)

                'por alguma razao precisa disso para ficar correto na tela durante a execucao
                valoresXTexto.AdjustAlignment(db)

            Next

            'zoom no grafico desenhado
            ZoomView(New Point2d(-5000, -5000), New Point2d((diaFinal - diaInicial) * distorcaoEixoX + 5000, maiorValorY + 5000))
            tr.Commit()
        End Using

    End Sub
    Private Sub AddToModelSpace(ent As Entity, db As Database)
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim modelSpace As BlockTableRecord = TryCast(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)

            modelSpace.AppendEntity(ent)
            tr.AddNewlyCreatedDBObject(ent, True)
            tr.Commit()
        End Using
    End Sub
    Private Sub CriarLayer(layerNome As String, lineTypeNome As String, indexColor As Integer)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        Using tr As Transaction = db.TransactionManager.StartTransaction()

            'abre como leitura 
            Dim layTable As LayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForWrite)

            Dim layTableRec As New LayerTableRecord With {
                .Name = layerNome
            }
            'se o layer nao existe cria um
            If layTable.Has(layTableRec.Name) = False Then
                '' adicionar a tabela de layers e a transação
                layTable.Add(layTableRec)
                tr.AddNewlyCreatedDBObject(layTableRec, True)
            Else
                layTableRec = tr.GetObject(layTable(layerNome), OpenMode.ForWrite)
            End If

            'carregar linetype , verifica se ja esta carregado
            Dim lineTypeTable = TryCast(tr.GetObject(db.LinetypeTableId, OpenMode.ForRead), LinetypeTable)
            If Not lineTypeTable.Has(lineTypeNome) Then
                db.LoadLineTypeFile(lineTypeNome, "acad.lin")
            End If

            Dim cor As Color
            cor = Color.FromColorIndex(ColorMethod.ByAci, indexColor)

            'muda o linetype
            layTableRec.LinetypeObjectId = lineTypeTable(lineTypeNome)
            'altera a cor
            layTableRec.Color = cor

            tr.Commit()
        End Using
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

        'converter o tipo dos dados, necessario para usar o compute
        dt.Columns(2).DataType = System.Type.GetType("System.Int32")
        dt.Columns(3).DataType = System.Type.GetType("System.Int32")
        dt.Columns(4).DataType = System.Type.GetType("System.Int32")
        dt.Columns(5).DataType = System.Type.GetType("System.Int32")
        dt.Columns(6).DataType = System.Type.GetType("System.Int32")

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

    Private Sub Abrir_Click(sender As Object, e As EventArgs) Handles Abrir.Click
        Dim AbrirComo As OpenFileDialog = New OpenFileDialog()
        Dim caminho As DialogResult

        Dim Arquivo As String

        AbrirComo.Title = "Abrir como"
        AbrirComo.Filter = "Arquivos de texto (*.csv)|*.csv"

        caminho = AbrirComo.ShowDialog
        Arquivo = AbrirComo.FileName

        'coloca o caminho do arqvuio no textbox
        TextBox_Caminho.Text = Arquivo
    End Sub

    Private Sub Button_Gerar_Click(sender As Object, e As EventArgs) Handles Button_Gerar.Click
        Me.Close()
        Plotar()
    End Sub
End Class