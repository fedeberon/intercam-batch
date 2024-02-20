Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports SocCam_Mantenimiento

<TestClass()> Public Class UnitTest1

    <TestMethod()> Public Sub CreacionCuotasSociales()
        Dim argParse As New ArgsParser

        Dim dbg As New List(Of String)
        dbg.Add("--crear-cuotas-sociales")
        dbg.Add("homologacion")
        dbg.Add($"ccb")
        dbg.Add($"p{1}")
        dbg.Add($"a{2022}")

        argParse.SW_DEBUG = True
        argParse.ParseArguments(Nothing, dbg)

        Dim ejecutor As New Executor
        ejecutor.Silent = True

        Dim homologacion As Boolean = argParse.PARAM_HOMOLOGACION

        Debug.Print(argParse.ToString)
        Assert.AreEqual(True, ejecutor.GenerarCuotasSocios(argParse.PARAM_PERIODO_CUOTA_SOCIAL - 1, argParse.PARAM_ANIO_CUOTA_SOCIAL, homologacion, argParse.SW_CUOTA_EXTRA_IMPORTE, argParse.SW_CUOTA_EXTRA_SECTORES, argParse.SW_CUOTA_OMITIR_SOCIOS_COFRE, argParse.SW_SENDMAIL))
    End Sub

End Class