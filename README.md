# imagbpmac
Para generar imágenes pop-up en PowerBI con "simple image", se requiere que las imágenes sean públicas en una url.
# Macro de la Primera Prueba:
Macro de la Primera Prueba
Option Explicit
Sub Combinar()
    'Dim shtHoja1 As Worksheet

    Dim strBP_Titulo As String
    Dim strCentroMAC As String
    Dim strBP_1Problema As String
    Dim strBP_2aObjetivo As String
    Dim strBP_Persona As String
    Dim strBP_2bComo As String
    Dim strBP_3aInvolucrados As String
    Dim strBP_4aBeneficiados As String
    Dim strBP_5Inversion As String
    Dim strBP_6Resultados As String
    Dim strCanal As String
    Dim strElemento As String
    Dim strConductor As String
    Dim strProceso As String
    Dim strAprendizaje As String
    Dim PrimeraFila As Long

    Dim objPPT As Object
    Dim objPres As Object
    Dim objSld As Object
    Dim objShp As Object

    'Set shtHoja1 = Worksheets("BDBPAC")
    Set objPPT = CreateObject("Powerpoint.Application")
    objPPT.Visible = True

    Set objPres = objPPT.presentations.Open("D:\Documentos\Trabajo\Chambas\2025 PCM\Adriana Raygada (Buenas Practicas)\Prueba\Prueba.pptx")
    objPres.SaveAs "D:\Documentos\Trabajo\Chambas\2025 PCM\Adriana Raygada (Buenas Practicas)\Prueba\PruebaOK.pptx"

    PrimeraFila = 2
 
  Do While Worksheets("BDBPAC").Cells(PrimeraFila, 1) <> ""
        strBP_Titulo = Worksheets("BDBPAC").Cells(PrimeraFila, 1)
        strCentroMAC = Worksheets("BDBPAC").Cells(PrimeraFila, 2)
        strBP_1Problema = Worksheets("BDBPAC").Cells(PrimeraFila, 3)
        strBP_2aObjetivo = Worksheets("BDBPAC").Cells(PrimeraFila, 4)
        strBP_Persona = Worksheets("BDBPAC").Cells(PrimeraFila, 5)
        strBP_2bComo = Worksheets("BDBPAC").Cells(PrimeraFila, 6)
        strBP_3aInvolucrados = Worksheets("BDBPAC").Cells(PrimeraFila, 7)
        strBP_4aBeneficiados = Worksheets("BDBPAC").Cells(PrimeraFila, 8)
        strBP_5Inversion = Worksheets("BDBPAC").Cells(PrimeraFila, 9)
        strBP_6Resultados = Worksheets("BDBPAC").Cells(PrimeraFila, 10)
        strCanal = Worksheets("BDBPAC").Cells(PrimeraFila, 11)
        strElemento = Worksheets("BDBPAC").Cells(PrimeraFila, 12)
        strConductor = Worksheets("BDBPAC").Cells(PrimeraFila, 13)
        strProceso = Worksheets("BDBPAC").Cells(PrimeraFila, 14)
        strAprendizaje = Worksheets("BDBPAC").Cells(PrimeraFila, 15)
        Set objSld = objPres.slides(1).Duplicate
        
        For Each objShp In objSld.Shapes
            If objShp.HasTextFrame Then
                If objShp.TextFrame.HasText Then
                    objShp.TextFrame.TextRange.Replace "<BP_Titulo>", strBP_Titulo
                    objShp.TextFrame.TextRange.Replace "<CentroMAC>", strCentroMAC
                    objShp.TextFrame.TextRange.Replace "<BP_1Problema>", strBP_1Problema
                    objShp.TextFrame.TextRange.Replace "<BP_2aObjetivo>", strBP_2aObjetivo
                    objShp.TextFrame.TextRange.Replace "<BP_Persona>", strBP_Persona
                    objShp.TextFrame.TextRange.Replace "<BP_2bComo>", strBP_2bComo
                    objShp.TextFrame.TextRange.Replace "<BP_3aInvolucrados>", strBP_3aInvolucrados
                    objShp.TextFrame.TextRange.Replace "<BP_4aBeneficiados>", strBP_4aBeneficiados
                    objShp.TextFrame.TextRange.Replace "<BP_5Inversion>", strBP_5Inversion
                    objShp.TextFrame.TextRange.Replace "<BP_6Resultados>", strBP_6Resultados
                    objShp.TextFrame.TextRange.Replace "<Canal>", strCanal
                    objShp.TextFrame.TextRange.Replace "<Elemento>", strElemento
                    objShp.TextFrame.TextRange.Replace "<Conductor>", strConductor
                    objShp.TextFrame.TextRange.Replace "<Proceso>", strProceso
                    objShp.TextFrame.TextRange.Replace "<Aprendizaje>", strAprendizaje
                End If
            End If
        Next
        PrimeraFila = PrimeraFila + 1
    Loop

    objPres.slides(1).Delete
    objPres.Save
End Sub
