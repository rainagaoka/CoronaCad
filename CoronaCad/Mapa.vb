Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports acaApp = Autodesk.AutoCAD.ApplicationServices.Application
Public Class Mapa
    <CommandMethod("CoronaCad")>
    Public Sub CoronaCad()

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim pr As PromptResult = ed.GetString(vbLf & "Nome do layer: ")

        If pr.Status = PromptStatus.OK Then
            Dim ents As ObjectIdCollection = ObjetosLayer(pr.StringResult)
            ed.WriteMessage(vbLf & "Found {0} entit{1} on layer {2}", ents.Count, (If(ents.Count = 1, "y", "ies")), pr.StringResult)
        End If
    End Sub

    Private Function ObjetosLayer(nomeLayer As String) As ObjectIdCollection
        Dim doc = acaApp.DocumentManager.MdiActiveDocument
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
End Class
