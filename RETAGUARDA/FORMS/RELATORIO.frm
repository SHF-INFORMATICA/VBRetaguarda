VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Begin VB.Form frmRELATORIO10 
   Caption         =   "Relatório"
   ClientHeight    =   8595
   ClientLeft      =   960
   ClientTop       =   2025
   ClientWidth     =   13035
   Icon            =   "RELATORIO.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   13035
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cmd 
      Left            =   3840
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   13035
      DesignHeight    =   8595
   End
   Begin CRVIEWERLibCtl.CRViewer Rel10 
      Height          =   9135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13095
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmRELATORIO10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"417E3D7B026A"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Form"
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS

   Dim TabRel      As New ADODB.Recordset
   Dim Caminho_Rel As String


Dim crxApp As New CRAXDRT.Application
Dim crxRpt As CRAXDRT.report
Dim crxTables As CRAXDRT.DatabaseTables
Dim crxTable As CRAXDRT.DatabaseTable
Dim crxSubreportObject As CRAXDRT.SubreportObject
Dim crxSubReport As CRAXDRT.report
Dim crxSections As CRAXDRT.Sections
Dim crxSection As CRAXDRT.Section

   Me.Top = 0
   Me.Left = 5

   'Me.Width = 12100
   'Me.Height = 8900

'==========================================
'==========================================
   Caminho_Rel = PATH_REL

   If TabRel.State = 1 Then _
      TabRel.Close

   SQL = "select * from IMPREL WITH (NOLOCK)"
   SQL = SQL & " where relatorio = '" & Trim(Nome_Relatorio) & "'"
   TabRel.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabRel.EOF Then _
      If Not IsNull(TabRel.Fields("caminho").Value) Then _
         If Trim(TabRel.Fields("caminho").Value) <> "" Then _
            Caminho_Rel = Trim(TabRel.Fields("caminho").Value)
   If TabRel.State = 1 Then _
      TabRel.Close

   Set crxReport = crxApplication.OpenReport(Caminho_Rel & Nome_Relatorio)

'crxReport.Database.LogOffServer "crdb_odbc.dll", Crystaldsn, Crystaldsq, Crystaluid, Crystalpwd
'crxApplication.LogOffServer "crdb_odbc.dll", Crystaldsn, Crystaldsq, Crystaluid, Crystalpwd

'==========================================
'atribui os parametros declarados aos objetos relacionados
'Dim crParameterDiscreteValue     As ParameterDiscreteValue
Dim crParameterFieldDefinitions  As ParameterFieldDefinitions
Dim crParameterFieldLocation     As ParameterFieldDefinition
Dim crParameterValues            As ParameterValues

'=====================
'Dim crParameterFields As New CrystalDecisions.Shared.ParameterFields()
'Dim crParameterField As New CrystalDecisions.Shared.ParameterField()
'Dim crParameterDiscreteValue As New CrystalDecisions.Shared.ParameterDiscreteValue()
'Dim crParameterValue As New CrystalDecisions.Shared.ParameterValues()
'Dim crParameterValues As New CrystalDecisions.Shared.ParameterValues()

'CrystalReportViewer1.ReportSource = "C:\Inetpub\wwwroot\itinerarypro\Itinerary.rpt"

'crParameterFields = CrystalReportViewer1.ParameterFieldInfo

'crParameterField = crParameterFields["@Resnumber"]

'crParameterValues = crParameterField.CurrentValues 'Set the current values for the parameter field
'crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue()
'crParameterDiscreteValue.Value = 101165
'crParameterValues.Add (crParameterDiscreteValue)

'CrystalReportViewer1.ParameterFieldInfo = crParameterFields

'=====================
' Pega a coleção de parametros do relatorio
'crParameterFieldDefinitions = CR.DataDefinition.ParameterFields
'==========================================

crxReport.EnableParameterPrompting = False
'crxReport.ParameterFields(0).ParameterType = "Percentual; " & txtperc.Text & " ;True"
'crxReport.ParameterFields(0).AddCurrentValue (1)
'SQL = "10"
'crxReport.ParameterFields(0).AddCurrentValue (1)


   crxReport.Database.Tables(1).SetLogOnInfo Crystaldsn, Crystaldsq, Crystaluid, Crystalpwd
   crxReport.DiscardSavedData

'MsgBox FORMULA_REL

'crxReport.ParameterFields(0) = "dtini;" & "vaca" & ";True"

   crxReport.RecordSelectionFormula = FORMULA_REL

Set crxTables = crxReport.Database.Tables
For Each crxTable In crxTables
    With crxTable
         .Location = .Name
    End With
Next

Set crxSections = crxReport.Sections

For i = 1 To crxSections.Count
    Set crxSection = crxSections(i)
    
    For j = 1 To crxSection.ReportObjects.Count
    
        If crxSection.ReportObjects(j).Kind = crSubreportObject Then
            Set crxSubreportObject = crxSection.ReportObjects(j)
            
            'Open the subreport, and treat like any other report
            Set crxSubReport = crxSubreportObject.OpenSubreport
            Set crxTables = crxSubReport.Database.Tables
            
            For Each crxTable In crxTables
               crxTable.SetLogOnInfo Crystaldsn, Crystaldsq, Crystaluid, Crystalpwd
               ' With crxTable
               '
               '     .SetLogOnInfo strServerOrDSNName, _
               '         strDBNameOrPath, strUserID, strPassword
                   crxTable.Location = "." & Name
               ' End With
            Next
        End If
    Next j
Next i
'========================
   'If Trim(UCase(NOME_BANCO_DADOS)) <> "MEGASIM" Then
   '   For Each crxTable In crxReport.Database.Tables
   '      crxTable.SetLogOnInfo Crystaldsn, Crystaldsq, Crystaluid, Crystalpwd
   '      scrxTableName = crxTable.Name
   '      crxTable.Location = Crystaldsq & "." & scrxTableName
   '      'crxTable.Location = "." & scrxTableName
   '      If Not crxTable.TestConnectivity Then
   '         '<can't connect error processing>
   '         Exit For
   '      End If
   '   Next
   'End If

   Rel10.Zoom 75
   Rel10.DisplayGroupTree = True

   Rel10.ReportSource = crxReport
   Rel10.ViewReport
'==========================================

   'crvRelatorio.Width = ScaleWidth
   'crvRelatorio.Height = ScaleHeight

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description & "  " & PATH_REL & Nome_Relatorio, Me.Name, "Form_Load"
End Sub

'##ModelId=417E3D7C0095
Private Sub Form_Resize()
   'Rel10.Width = ScaleWidth
   'Rel10.Height = ScaleHeight
End Sub

'##ModelId=417E3D7C0096
Private Sub Form_Unload(Cancel As Integer)
   Set crxReport = Nothing
   Set crxApplication = Nothing
   Set crxSubReport = Nothing
End Sub

'##ModelId=417E3D7C00A6
Private Sub imgFecharOn_Click()
   Unload Me
End Sub

'##ModelId=417E3D7C00B5
Private Sub imgMinimizarOn_Click()
   Me.WindowState = vbMinimized
   Me.SetFocus
End Sub
