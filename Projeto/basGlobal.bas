Attribute VB_Name = "basGlobal"
Option Explicit

Public CaminhoApp As String

' ??? - deu problema aqui em função da nova estrutura
' de frames do frmInterface
Public Function PegaBanco(Formulário As Form) As String
    Dim DC As Object
    For Each DC In Formulário
        If TypeOf DC Is Data Then
            DC.DatabaseName = CaminhoApp & "..\Banco\Projetos.mdb"
            DC.Refresh
        End If
    Next DC
End Function

Public Sub Status(TextoDeAjuda As String)
    frmAmbiente.StatusBar1.Panels(1).Text = TextoDeAjuda
End Sub
