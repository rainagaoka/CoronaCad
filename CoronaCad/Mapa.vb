Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports acaApp = Autodesk.AutoCAD.ApplicationServices.Application
Public Class Mapa
    <CommandMethod("CoronaCad")>
    Public Sub CoronaCad()

        Dim doc = acaApp.DocumentManager.MdiActiveDocument

        Using tr As Transaction = doc.TransactionManager.StartOpenCloseTransaction

        End Using
    End Sub
End Class
