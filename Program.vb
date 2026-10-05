Option Explicit On
Option Strict On
Imports CATIAClassLibrary
Module Program

    Sub Main()


        Console.WriteLine("> Starting Process...")

        ' Catia
        Dim CATIAsession As New CatiaSession
        If Not CATIAsession.IsReady Then
            MsgBox(CATIAsession.Description)
            Exit Sub
        End If
        Dim oProduct As ProductStructureTypeLib.Product = CATIAsession.RootProduct
        CATIAsession.Application.DisplayFileAlerts = False



        ' Directorios y nombres
        Dim baseDir As String = "C:\Temp"
        Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
        Dim folderPath As String = IO.Path.Combine(baseDir, "Export_" & timestamp)
        Dim excelFileName As String = IO.Path.Combine(folderPath, "Reporte_" & timestamp & ".xlsx")
        If Not IO.Directory.Exists(folderPath) Then
            IO.Directory.CreateDirectory(folderPath)
        End If


        ' Extracción CATIA
        Dim oCatiaDataextractor As New CatiaDataExtractor
        Dim oDataTable As DataTable = oCatiaDataextractor.ExtractData(oProduct, folderPath)



    End Sub

End Module
