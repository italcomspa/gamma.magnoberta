VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLSBO_REGDOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum tsErrInitDoc
    tsOK = 0
    tsErrInitializeBO = 1
    tsErrOpenEmptyRecordsets = 2
    tsErrDocEliminazione = 3
    tsErrProgram = 99
End Enum

' enumeratore tipologia riga
Public Enum tsEnumImportDocumentiTipoRiga
    tsImportDocumentiArticoloMagazzino = 0
    tsImportDocumentiArticoloManuale = 1
    tsImportDocumentiRigaDescrittiva = 2
    tsImportDocumentiSpesa = 3
    tsImportDocumentiTestoFisso = 4
    tsImportDocumentiFaseLavorazione = 5
End Enum

Public Errore                       As tsErrInitDoc
Public ErrDescr                     As Variant

Public ActiveClass                  As CLSGB_ORDINE
Public ClsRegDoc                    As MGBO_REGDOCUMENTI.CLSMG_REGDOC

Private CodiceDitta                 As String
Private strSQL                      As String
Public Gcon_Connect                 As ADODB.Connection
Public ActiveInterface              As Cinterface

Private Pbol_EseguitaAddNew         As Boolean

Public NumregDocumento              As String
Public NumDocumento                 As String
Public ProgRigaDocumento            As Double
Public DataDocumento                As Date
Public CodiceDocCollegato            As String




Public Sub InizializzaBODocumenti(Optional ByVal OpenEmptyRecords As Boolean = True, _
                                   Optional ByVal TransactionON As Boolean = False, _
                                   Optional ByVal CodiceDocumento As String)
  On Error GoTo ErrTrap

  CodiceDitta = ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta

  If ClsRegDoc Is Nothing Then
      Set ClsRegDoc = New MGBO_REGDOCUMENTI.CLSMG_REGDOC
      Set ClsRegDoc.RegDocIn.ActiveInterface = ActiveInterface
      Set ClsRegDoc.RegDocIn.Connection = Gcon_Connect
      ClsRegDoc.RegDocIn.Ditta = CDec(CodiceDitta)
      ClsRegDoc.RegDocIn.Documento = CodiceDocumento
  Else
      ClsRegDoc.Terminate
  End If

  Errore = 0
  ErrDescr = ""
  
  With ClsRegDoc
  
'    .RegDocIn.BatchMode = True
    .RegDocIn.BatchMode = False
    .RegDocIn.Avanzamento = True
    
    .RegDocIn.GestDocCorClRif = False
    .RegDocIn.GestDocCorOrdDet = False
    .RegDocIn.GestDocCorpoBVM = False
    .RegDocIn.GestDocCorpoConf = False
    .RegDocIn.GestDocCorpoEst = False
    .RegDocIn.GestDocCorpoOrd = True
    .RegDocIn.GestDocCorpoPers = True
    .RegDocIn.GestDocCorpoProge = True
    .RegDocIn.GestDocCorpoProv = False
    .RegDocIn.GestDocCorpoRatei = False
    .RegDocIn.GestDocCorpoRif = True
    .RegDocIn.GestDocCorpoUbic = False
    .RegDocIn.GestDocCorPrel = False
    .RegDocIn.GestDocCorpoLot = True
    .RegDocIn.GestDocCorpoCoin = True
    
    .RegDocIn.GestDocTestata = True
    .RegDocIn.GestDocTotali = True
    .RegDocIn.TotaliSenzaContabilita = False
    
    .RegDocIn.GestDocDatiAcc = True
    .RegDocIn.GestDocTesAgenti = True
    .RegDocIn.GestDocTesOrd = True
    .RegDocIn.GestDocTesOrdAP = True
    .RegDocIn.GestDocTestaEst = True
    .RegDocIn.GestDocTestaPers = True
    .RegDocIn.GestDocTestaProge = True
    .RegDocIn.GestDocTestaProv = False
    .RegDocIn.GestDocTestaRatei = False
    .RegDocIn.GestDocTestaRif = True
    .RegDocIn.GestDocTestaRifFN = True
    
    If TransactionON Then
      .RegDocIn.Transaction = True
    End If
    Set .RegDocIn.Connection = ActiveInterface.Connection
    .Initialize
    
    If .Stato <> 0 Then
      Errore = tsErrInitializeBO
      ErrDescr = .StatoDesc
    End If
    
    If OpenEmptyRecords Then
      .OpenAllEmptyRecordsets
    
      If .Stato <> 0 Then
          Errore = tsErrOpenEmptyRecordsets
          ErrDescr = .StatoDesc
      End If
    End If
  
  End With

  Exit Sub

ErrTrap:
  Errore = tsErrProgram
  ErrDescr = Err.Description
End Sub


Public Sub StampaDocumento(ByVal rCodDoc As String, _
                            ByVal rNumReg As String)

On Error GoTo ErrEliminaDocumento

  ClsRegDoc.OpenDocTestata "DO11_NUMREG_CO99 = '" & rNumReg & "'"
  If ClsRegDoc.rstDocTestata.RecordCount > 0 Then
      ClsRegDoc.StampaDocumento
  End If

  Exit Sub
ErrEliminaDocumento:
    Errore = tsErrProgram
    ErrDescr = Err.Description
End Sub

Public Sub EliminaDocumento(ByVal rCodDoc As String, _
                            ByVal rNumReg As String)

On Error GoTo ErrEliminaDocumento

  ClsRegDoc.OpenDocTestata "DO11_NUMREG_CO99 = '" & rNumReg & "'"
  If ClsRegDoc.rstDocTestata.RecordCount > 0 Then
      ClsRegDoc.DeleteDocTestata
      If ClsRegDoc.Stato <> 0 Then
        ErrDescr = ClsRegDoc.StatoDesc & vbNewLine & _
                 "Eliminazione doc. " & rCodDoc & " n.r." & rNumReg
        Errore = tsErrDocEliminazione
      End If
  End If

  Exit Sub
ErrEliminaDocumento:
    Errore = tsErrProgram
    ErrDescr = Err.Description
End Sub

Public Sub RilasciaBODocumenti()
    On Error Resume Next
    If Not ClsRegDoc Is Nothing Then
        ClsRegDoc.Terminate
        Set ClsRegDoc.RegDocIn.ActiveInterface = Nothing
        Set ClsRegDoc.RegDocIn.Connection = Nothing
        Set ClsRegDoc.RegDocIn = Nothing
        Set ClsRegDoc = Nothing
    End If

    Err.Clear
End Sub


Public Function RegistraDocumento()

On Error GoTo Err_RegistraDocumento

   Errore = 0
   ErrDescr = ""

   If Not ClsRegDoc Is Nothing Then
       ClsRegDoc.UpdateDocTestata , True
        If NVL(CodiceDocCollegato, "") <> "" Then
                GeneraDocumentoCollegato (CodiceDocCollegato)
        End If
   End If
   
   NumregDocumento = ClsRegDoc.rstDocTestata("DO11_NUMREG_CO99").value
   NumDocumento = ClsRegDoc.rstDocTestata("DO11_NUMDOC").value
   DataDocumento = ClsRegDoc.rstDocTestata("DO11_DATADOC").value
   
   
   
   
   Errore = ClsRegDoc.Stato
   ErrDescr = ClsRegDoc.StatoDesc

Exit Function
Err_RegistraDocumento:
   Errore = tsErrProgram
   ErrDescr = Err.Description
End Function


Public Sub UpdateTestataDocumento(ByVal rCodDoc As String, _
                                  ByVal rNumReg As String)

On Error GoTo ErrTrap

  With ClsRegDoc
    .OpenDocTestata " DO11_NUMREG_CO99 = '" & rNumReg & "'"
    .ModifyDocTestata rNumReg, True
    
'    .UpdateDocTestata
  End With
  
'  If ClsRegDoc.rstDocTestata.RecordCount > 0 Then
'      ClsRegDoc.DeleteDocTestata
'      If ClsRegDoc.Stato <> 0 Then
'        ErrDescr = ClsRegDoc.StatoDesc & vbNewLine & _
'                 "Eliminazione doc. " & rCodDoc & " n.r." & rNumReg
'        Errore = tsErrDocEliminazione
'      End If
'  End If

  Exit Sub
ErrTrap:
    Errore = tsErrProgram
    ErrDescr = Err.Description
End Sub

'
' Aggiunta riga di testata documento
'
Public Function AddNewTestataDocumento(Optional ByVal CodiceClienteFornitore As Variant, Optional ByVal CodiceAgente As Variant, _
                                       Optional ByVal Cambio As Variant, Optional ByVal CausaleMagazzino As Variant, _
                                       Optional ByVal CodiceDeposito As Variant, Optional ByVal CodiceAliquotaIVA As Variant, _
                                       Optional ByVal CodicePagamento As Variant, Optional ByVal DataCambio As Variant, _
                                       Optional ByVal DataDocumento As Variant, Optional ByVal DataRegistrazione As Variant, _
                                       Optional ByVal DocumentoBis As Boolean = False, Optional ByVal PartitaBis As Boolean = False, _
                                       Optional ByVal SpeseBolli As Boolean = False, Optional ByVal SpeseIncasso As Boolean = False, _
                                       Optional ByVal NumeroDocumento As Variant, Optional ByVal NumeroDocumentoOrigine As Variant, _
                                       Optional ByVal NumeroPartita As Variant, Optional ByVal SezionaleDocumento As Variant, _
                                       Optional ByVal SezionalePartita As Variant, Optional ByVal CodiceValuta As Variant, _
                                       Optional ByVal DestinatarioPrevalente As Variant, _
                                       Optional ByVal NumeroDocumentoRiferimento As Variant, _
                                       Optional ByVal NotaDocumentoRiferimento As String, _
                                       Optional ByVal SezionaleDocumentoRiferimento As String, _
                                       Optional ByVal DataDocumentoRiferimento As Variant, _
                                       Optional ByVal RifVostroDoc As Variant, _
                                       Optional ByVal NumVostroDoc As Variant, _
                                       Optional ByVal RifNsDoc As Variant, _
                                       Optional ByVal NumNsDoc As Variant, _
                                       Optional ByVal DataNsDoc As Variant, _
                                       Optional ByVal DataConfermaNsDoc As Variant, _
                                       Optional ByVal CIG As Variant, _
                                       Optional ByVal CUP As Variant, _
                                       Optional ByVal Commessa As Variant, Optional ByVal SottoCommessa As Variant, _
                                       Optional ByVal Vettore1 As Variant, Optional ByVal Responsabile As Variant, Optional ByVal RagioneSocialeAgente As Variant, Optional ByVal AlfPers1Cig As Variant, Optional ByVal AlfPers2Cup As Variant)

    On Error GoTo Err_AddNewTestataDocumento

   Errore = 0
   ErrDescr = ""

   With ClsRegDoc
        '.AddNewDocTestata  False, True, True
        'problema trinchero
        .AddNewDocTestata False, False, True
        
        .AddNewDocTotali False
        .AddNewDocDatiAcc False
        
        'fine modifiche problema trinchero
        
        Pbol_EseguitaAddNew = True

        If Not IsMissing(CodiceAliquotaIVA) Then
            .CambiaDocTestataAliva CodiceAliquotaIVA
        End If
        
        If Not IsMissing(CIG) Then
            '.CambiaDocTestataCIG CIG
           .rstDocTestata.Fields("DO11_CIG").value = CIG
        End If
        
         If Not IsMissing(CUP) Then
            '.CambiaDocTestataCUP CUP
            .rstDocTestata.Fields("DO11_CUP").value = CUP
        End If
                       
        If Not IsMissing(Cambio) Then
            .CambiaDocTestataCambio Cambio
        End If

        If Not IsMissing(CausaleMagazzino) Then
            .CambiaDocTestataCausMag CausaleMagazzino
        End If

        If Not IsMissing(CodiceClienteFornitore) Then
            .CambiaDocTestataCliFor CodiceClienteFornitore, True
        End If

        If Not IsMissing(CodiceAgente) Then
            .CambiaDocTesAgentiAgente CodiceAgente
        End If

        If Not IsMissing(CodiceDeposito) Then
        'MsgBox CodiceDeposito
            .CambiaDocTestataCodDep CodiceDeposito
        End If

        If Not IsMissing(CodicePagamento) Then
            .CambiaDocTestataCodPag CodicePagamento
        End If

        If Not IsMissing(DataCambio) Then
            .CambiaDocTestataDataCambio DataCambio
        End If

        If Not IsMissing(DataDocumento) Then
            .CambiaDocTestataDataDoc DataDocumento
        End If

        If Not IsMissing(DataRegistrazione) Then
            .CambiaDocTestataDataReg DataRegistrazione
        End If

        If Not IsMissing(DocumentoBis) Then
            .CambiaDocTestataFlgDocBis DocumentoBis
        End If

        If Not IsMissing(PartitaBis) Then
            .CambiaDocTestataFlgPartBis PartitaBis
        End If

        If Not IsMissing(SpeseBolli) Then
            .CambiaDocTestataFlgSpBolli SpeseBolli
        End If

        If Not IsMissing(SpeseIncasso) Then
            .CambiaDocTestataFlgSpincas SpeseIncasso
        End If

        If Not IsMissing(NumeroDocumento) Then
            .CambiaDocTestataNumDoc NumeroDocumento
        End If

        If Not IsMissing(NumeroDocumentoOrigine) Then
          If NVL(NumeroDocumentoOrigine, "") <> "" Then
            .CambiaDocTestataNumDocOrig NumeroDocumentoOrigine
          End If
        End If

        If Not IsMissing(NumeroPartita) Then
            .CambiaDocTestataNumPart NumeroPartita
        End If

        If Not IsMissing(SezionaleDocumento) Then
            .CambiaDocTestataSezDoc SezionaleDocumento
        End If

        If Not IsMissing(SezionalePartita) Then
            .CambiaDocTestataSezPart SezionalePartita
        End If

        If Not IsMissing(CodiceValuta) Then
            .CambiaDocTestataValuta CodiceValuta
        End If

        If Not IsMissing(DestinatarioPrevalente) Then
          .CambiaDocTestataCliForDest Trim(DestinatarioPrevalente)
'          .rstDocTestata.Fields("DO11_CLIFORDEST").Value = Trim(DestinatarioPrevalente)
        End If
        
'        If Not IsMissing(DepositoCollegato) Then
'          If NVL(DepositoCollegato, "") <> "" Then
'          MsgBox DepositoCollegato
'            .CambiaDocTestataCodDepCol DepositoCollegato
'          End If
'        End If
        
        If Not IsMissing(Vettore1) Then
          .CambiaDocDatiAccVettore1 Vettore1
        End If
        
        
        
        .AddNewDocTestaRif

        
        If Not IsMissing(NumeroDocumentoRiferimento) Then
          If NVL(NumeroDocumentoRiferimento, "") <> "" Then
            .rstDocTestaRif.Fields("DO12_NUMNSDOC").value = NumeroDocumentoRiferimento
          End If
        End If
        If Not IsMissing(DataDocumentoRiferimento) Then
          If NVL(DataDocumentoRiferimento, "") <> "" Then
            .rstDocTestaRif.Fields("DO12_DATANSDOC").value = DataDocumentoRiferimento
          End If
        End If
        If Not IsMissing(SezionaleDocumentoRiferimento) Then
          .rstDocTestaRif.Fields("DO12_SEZNSDOC").value = SezionaleDocumentoRiferimento
        End If
        If Not IsMissing(NotaDocumentoRiferimento) Then
          .rstDocTestaRif.Fields("DO12_RIFNSDOC").value = NotaDocumentoRiferimento
        End If
        
        If Not IsMissing(RifVostroDoc) Then
          .rstDocTestaRif.Fields("DO12_RIFVSDOC").value = RifVostroDoc
        End If
        
        If Not IsMissing(NumVostroDoc) Then
          .rstDocTestaRif.Fields("DO12_NUMVSDOC").value = CDbl(Trim(NumVostroDoc))
        End If
        
        If Not IsMissing(RifNsDoc) Then
          .rstDocTestaRif.Fields("DO12_RIFNSDOC").value = RifNsDoc
        End If
        
        If Not IsMissing(NumNsDoc) Then
          .rstDocTestaRif.Fields("DO12_NUMNSDOC").value = CDbl(Trim(NumNsDoc))
        End If
        
        If Not IsMissing(DataNsDoc) Then
          .rstDocTestaRif.Fields("DO12_DATANSDOC").value = DataNsDoc
        End If
        
        If Not IsMissing(DataConfermaNsDoc) Then
          .rstDocTestaRif.Fields("DO12_DATANSCONF").value = DataConfermaNsDoc
        End If
        
        
         If Not IsMissing(AlfPers1Cig) Then
        
            If .rstDocTestaPers.RecordCount = 0 Then
                Call .AddNewDocTestaPers
            End If
            .rstDocTestaPers.Fields("DO17_ALFPERS1").value = AlfPers1Cig
            .UpdateDocTestaPers
        End If
        
         If Not IsMissing(AlfPers2Cup) Then
        
            If .rstDocTestaPers.RecordCount = 0 Then
                Call .AddNewDocTestaPers
            End If
            .rstDocTestaPers.Fields("DO17_ALFPERS2").value = AlfPers2Cup
            .UpdateDocTestaPers
        End If
         
        
        .UpdateDocTestaRif
        
        If Not IsMissing(Responsabile) Then
        
            If .rstDocTestaEst.RecordCount = 0 Then
                Call .AddNewDocTestaEst
            End If
            .rstDocTestaEst.Fields("DO20_ALFST4").value = Responsabile
            .UpdateDocTestaEst
        End If
        
         If Not IsMissing(RagioneSocialeAgente) Then
        
            If .rstDocTestaEst.RecordCount = 0 Then
                Call .AddNewDocTestaEst
            End If
            .rstDocTestaEst.Fields("DO20_ALFST3").value = RagioneSocialeAgente
            .UpdateDocTestaEst
        End If
        
        
         If Not IsMissing(Commessa) Then
            If .rstDocTestaProge.RecordCount = 0 Then
                Call .AddNewDocTestaProge
            End If
        
            .rstDocTestaProge.Fields("DO24_CODCOMMESSA").value = Commessa
            If Not IsMissing(SottoCommessa) Then
              
              .rstDocTestaProge.Fields("DO41_CODSCOMMESSA").value = SottoCommessa
            End If
            .UpdateDocTestaProge
         End If
        
        
        
    End With

   Errore = ClsRegDoc.Stato
   ErrDescr = ClsRegDoc.StatoDesc

  Exit Function

Err_AddNewTestataDocumento:
  Errore = tsErrProgram
  ErrDescr = Err.Description
  
End Function


Public Sub UpdateCorpoDoc(ByVal rNumReg As String, ByVal rProgRiga As String, _
                            ByVal TipoRiga As tsEnumImportDocumentiTipoRiga, Optional ByVal CodiceArticolo As Variant, Optional ByVal CodiceAliquotaIVA As Variant, _
                            Optional ByVal CodiceVarianteArticolo As Variant, Optional ByVal CodiceDeposito As Variant, _
                            Optional ByVal CausaleMagazzino As Variant, _
                            Optional ByVal CodiceSpesa As Variant, Optional ByVal CodiceFaseLavorazione As Variant, _
                            Optional ByVal MaggiorazioneImporto As Variant, Optional ByVal PercentualeMaggiorazione1 As Variant, _
                            Optional ByVal PercentualeMaggiorazione2 As Variant, Optional ByVal Prezzo As Variant, _
                            Optional ByVal NumColli As Variant, Optional ByVal Quantita As Variant, Optional ByVal ScontoImporto As Variant, _
                            Optional ByVal PercentualeSconto1 As Variant, Optional ByVal PercentualeSconto2 As Variant, _
                            Optional ByVal PercentualeSconto3 As Variant, Optional ByVal PercentualeSconto4 As Variant, _
                            Optional ByVal PercentualeSconto5 As Variant, Optional ByVal PercentualeSconto6 As Variant, _
                            Optional ByVal CodiceTestoFisso As Variant, Optional ByVal DesRiga As Variant, _
                            Optional ByVal PercentualeProvvigioni As Variant, Optional ByVal CodiceLotto As Variant, _
                            Optional ByVal ScadenzaLotto As Variant, Optional ByVal CodiceUbicazione As Variant, _
                            Optional ByVal RiferimentoNumreg As Variant, Optional ByVal RiferimentoProgriga As Variant, _
                            Optional ByVal RiferimentoTipoDoc As Variant, Optional ByVal RiferimentoSottoTipoDoc As Variant, _
                            Optional ByVal RiferimentoNumDoc As Variant, Optional ByVal RiferimentoSezDoc As Variant, _
                            Optional ByVal RiferimentoDocBis As Variant, Optional ByVal RiferimentoDataDoc As Variant, _
                            Optional ByVal RiferimentoProgrigaMaga As Variant, Optional ByVal RiferimentoProgrigaLotto As Variant, _
                            Optional ByVal VDS As Variant, Optional ByVal CDC As Variant, Optional ByVal Commessa As Variant, _
                            Optional ByVal UM1 As Variant, Optional ByVal UM2 As Variant, Optional ByVal FatConv As Variant, _
                            Optional ByVal PZConf As Variant, Optional ByVal PesoL As Variant, Optional ByVal PesoN As Variant, _
                            Optional ByVal Tara As Variant, Optional ByVal SottoCommessa As Variant, Optional ByVal StatoRigaOfferta As Variant)


On Error GoTo ErrTrap
  Dim strFiltro As String
  
  strFiltro = "DO30_NUMREG_CO99 = " & rNumReg
  strFiltro = strFiltro & " AND DO30_PROGRIGA = " & CDbl(rProgRiga)
  
  Errore = 0
  ErrDescr = ""

  With ClsRegDoc
'
'    ' If Not Pbol_EseguitaAddNew Then
'     ' problema trinchero
'    .AddNewDocCorpo False, , , True
'     'End If
'     'Pbol_EseguitaAddNew = False
    
    .OpenDocCorpo strFiltro
    
    If Not IsMissing(CodiceAliquotaIVA) Then
      .CambiaDocCorpoAliva CodiceAliquotaIVA
    End If
    
    Select Case TipoRiga
       Case tsImportDocumentiArticoloMagazzino
           .rstDocCorpo("DO30_INDTIPORIGA").value = 0
       Case tsImportDocumentiArticoloManuale
           .rstDocCorpo("DO30_INDTIPORIGA").value = 1
       Case tsImportDocumentiFaseLavorazione
           .rstDocCorpo("DO30_INDTIPORIGA").value = 10
       Case tsImportDocumentiRigaDescrittiva
           .rstDocCorpo("DO30_INDTIPORIGA").value = 2
       Case tsImportDocumentiSpesa
           .rstDocCorpo("DO30_INDTIPORIGA").value = 4
       Case tsImportDocumentiTestoFisso
           .rstDocCorpo("DO30_INDTIPORIGA").value = 6
    End Select
    
    If Not IsMissing(UM1) Then
      .rstDocCorpo("DO30_UM1").value = UM1
    End If
    
    If Not IsMissing(UM2) Then
      .rstDocCorpo("DO30_UM2").value = UM2
    End If
  
    If Not IsMissing(FatConv) Then
      .rstDocCorpo("DO30_FATCONV").value = FatConv
    End If
  
    If Not IsMissing(PZConf) Then
      .rstDocCorpo("DO30_PZCONF").value = PZConf
    End If
  
    If Not IsMissing(PesoL) Then
      .rstDocCorpo("DO30_PESOL").value = PesoL
    End If
    
    If Not IsMissing(PesoN) Then
      .rstDocCorpo("DO30_PESON").value = PesoN
    End If
  
    If Not IsMissing(Tara) Then
      .rstDocCorpo("DO30_TARA").value = Tara
    End If
  
'    If Not IsMissing(Capacita) Then
'      .rstDocCorpo("DO30_CAPACITA").Value = Capacita
'    End If
  
    If Not IsMissing(DesRiga) Then
      .rstDocCorpo("DO30_DESCART").value = DesRiga
    End If
    
    If Not IsMissing(CodiceArticolo) Then
      .CambiaDocCorpoCodart CodiceArticolo, CodiceVarianteArticolo, True
    End If
    
    If Not IsMissing(CodiceDeposito) Then
       .CambiaDocCorpoCodDep CodiceDeposito
    End If
    
    If Not IsMissing(CodiceSpesa) Then
       .CambiaDocCorpoCodSpesa CodiceSpesa
    End If
    
    If Not IsMissing(CodiceFaseLavorazione) Then
       .CambiaDocCorpoFase CodiceFaseLavorazione
    End If
    
    If Not IsMissing(MaggiorazioneImporto) Then
       .CambiaDocCorpoMagImp MaggiorazioneImporto
    End If
    
    If Not IsMissing(PercentualeMaggiorazione1) Then
       .CambiaDocCorpoMagPer1 PercentualeMaggiorazione1
    End If
    
    If Not IsMissing(PercentualeMaggiorazione2) Then
       .CambiaDocCorpoMagPer2 PercentualeMaggiorazione2
    End If
    
    If Not IsMissing(Prezzo) Then
       .CambiaDocCorpoPrezzo1 Prezzo
    End If
    
    If Not IsMissing(NumColli) Then
       .CambiaDocCorpoColli NumColli
    End If
    
    If Not IsMissing(Quantita) Then
       .CambiaDocCorpoQta1 Quantita, True, True, True
    End If
    
    If Not IsMissing(ScontoImporto) Then
       .CambiaDocCorpoScImp ScontoImporto
    End If
    
    If Not IsMissing(PercentualeSconto1) Then
       .CambiaDocCorpoScPer1 PercentualeSconto1
    End If
    
    If Not IsMissing(PercentualeSconto2) Then
       .CambiaDocCorpoScPer2 PercentualeSconto2
    End If
    
    If Not IsMissing(PercentualeSconto3) Then
       .CambiaDocCorpoScPer3 PercentualeSconto3
    End If
    
    If Not IsMissing(PercentualeSconto4) Then
       .CambiaDocCorpoScPer4 PercentualeSconto4
    End If
    
    If Not IsMissing(PercentualeSconto5) Then
       .CambiaDocCorpoScPer5 PercentualeSconto5
    End If
    
    If Not IsMissing(PercentualeSconto6) Then
       .CambiaDocCorpoScPer6 PercentualeSconto6
    End If
    
    If Not IsMissing(CodiceTestoFisso) Then
       .CambiaDocCorpoTestoFix CodiceTestoFisso
    End If
    
    If Not IsMissing(CausaleMagazzino) Then
      .rstDocCorpo.Fields("DO30_CAUSMAG_MG51").value = CausaleMagazzino
    End If
    
     
    .UpdateDocCorpo , True, , False
  
    'COAN
    If Not IsMissing(VDS) Or Not IsMissing(CDC) Then
    '          If NVL(VDS, "") <> "" Or NVL(CDC, "") <> "" Then
      If .rstDocCorpoCoin.RecordCount = 0 Then
          Call .AddNewDocCorpoCoin
      End If
  
      If IsMissing(CDC) Then
        .rstDocCorpoCoin.Fields("DO53_PDCDEST_PC01").value = ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsParCOIN.CodicePDC
        .rstDocCorpoCoin.Fields("DO53_CONTODEST_PC03").value = CDC
      End If
  
      If Not IsMissing(VDS) Then
        .rstDocCorpoCoin.Fields("DO53_PDCVDSDEST_PC01").value = ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsParCOIN.CodiceVdS
        .rstDocCorpoCoin.Fields("DO53_VDSDEST_PC03").value = VDS
      End If
  
      .UpdateDocCorpoCoin
  
    End If
  
  
    If Not IsMissing(Commessa) Then
      If .rstDocCorpoProge.RecordCount = 0 Then
          Call .AddNewDocCorpoProge
      End If
  
      .rstDocCorpoProge.Fields("DO41_CODCOMMESSA").value = Commessa
      If Not IsMissing(SottoCommessa) Then
        
        .rstDocCorpoCoin.Fields("DO41_CODSCOMMESSA").value = SottoCommessa
      End If
      .UpdateDocCorpoProge
    End If
    
'    If Not IsMissing(StatoRigaOfferta) Then
'    .AddNewDocCorpoPers False
'        .rstDocCorpoRif.Fields("DO35_DITTA_CG18").value = CodiceDitta
'        .rstDocCorpoRif.Fields("DO35_NUMREG_CO99").value = .rstDocCorpo("DO30_NUMREG_CO99")
'        .rstDocCorpoRif.Fields("DO35_PROGRIGA").value = .rstDocCorpo("DO30_PROGRIGA")
'        .rstDocCorpoRif.Fields("DO35_ALFPERS1").value = StatoRigaOfferta
'    .UpdateNewDocCorpoPers
'    End If
  
    If Not IsMissing(RiferimentoNumreg) And Not IsMissing(RiferimentoProgriga) Then
      If NVL(RiferimentoNumreg, "") <> "" And NVL(RiferimentoProgriga, "") <> "" Then
        'Inserimento riferimenti
        .AddNewDocCorpoRif False
  
        .rstDocCorpoRif.Fields("DO33_NUMREG_CO99").value = .rstDocCorpo("DO30_NUMREG_CO99")
        .rstDocCorpoRif.Fields("DO33_PROGRIGA").value = .rstDocCorpo("DO30_PROGRIGA")
        .rstDocCorpoRif.Fields("DO33_PROGRIF").value = 1
        
  
        If Not IsMissing(RiferimentoNumreg) Then
          .rstDocCorpoRif.Fields("DO33_NUMREGRIF_CO99").value = RiferimentoNumreg
        End If
  
        If Not IsMissing(RiferimentoProgriga) Then
          .rstDocCorpoRif.Fields("DO33_PROGRIGARIF_DO30").value = RiferimentoProgriga
        End If
  
        If Not IsMissing(RiferimentoTipoDoc) Then
          .rstDocCorpoRif.Fields("DO33_INDTIPODOC").value = RiferimentoTipoDoc
        End If
  
        If Not IsMissing(RiferimentoSottoTipoDoc) Then
          .rstDocCorpoRif.Fields("DO33_INDSTIPODOC").value = RiferimentoSottoTipoDoc
        End If
  
        If Not IsMissing(RiferimentoNumDoc) Then
          .rstDocCorpoRif.Fields("DO33_NUMNSDOC").value = RiferimentoNumDoc
        End If
  
        If Not IsMissing(RiferimentoSezDoc) Then
          .rstDocCorpoRif.Fields("DO33_SEZNSDOC").value = RiferimentoSezDoc
        End If
  
        If Not IsMissing(RiferimentoDocBis) Then
          .rstDocCorpoRif.Fields("DO33_FLGNSDOCBIS").value = RiferimentoDocBis
        End If
  
        If Not IsMissing(RiferimentoDataDoc) Then
          .rstDocCorpoRif.Fields("DO33_DATANSDOC").value = RiferimentoDataDoc
        End If
  
        .rstDocCorpoRif.Fields("DO33_INDQTAORDRES").value = 0
        .rstDocCorpoRif.Fields("DO33_FLGAGGQTATRASF").value = 1
        .rstDocCorpoRif.Fields("DO33_FLGSTORNOQTAVAL").value = 1
  
        If Not IsMissing(Quantita) Then
          .rstDocCorpoRif.Fields("DO33_QTA1MOV").value = .rstDocCorpo("DO30_QTA1").value
          .rstDocCorpoRif.Fields("DO33_QTA2MOV").value = .rstDocCorpo("DO30_QTA2").value
  '             .CambiaDocCorpoQta1 Quantita
        End If
        
        If Not IsMissing(NumColli) Then
           .rstDocCorpoRif.Fields("DO33_COLLIMOV").value = NumColli
        End If
  
        .UpdateDocCorpoRif
  
      End If
    End If
    
    .RicalcolaCostoRiga
    .RicalcolaPrezzo
    .RicalcolaImportoRiga
    .RicalcolaTotaliDaRigheCorpo , , , False
    
    
    .UpdateDocCorpo , True, , False
          
    NumregDocumento = .rstDocCorpo("DO30_NUMREG_CO99")
    ProgRigaDocumento = .rstDocCorpo("DO30_PROGRIGA")
  
  End With
  
  
  
  Exit Sub
ErrTrap:
    Errore = tsErrProgram
    ErrDescr = Err.Description
End Sub


'
' Aggiunta Riga del corpo documento
'
Public Function AddNewRigaDocumento(ByVal TipoRiga As tsEnumImportDocumentiTipoRiga, Optional ByVal CodiceArticolo As Variant, Optional ByVal CodiceAliquotaIVA As Variant, _
                                    Optional ByVal CodiceVarianteArticolo As Variant, Optional ByVal CodiceDeposito As Variant, _
                                    Optional ByVal CodiceSpesa As Variant, Optional ByVal CodiceFaseLavorazione As Variant, _
                                    Optional ByVal MaggiorazioneImporto As Variant, Optional ByVal PercentualeMaggiorazione1 As Variant, _
                                    Optional ByVal PercentualeMaggiorazione2 As Variant, Optional ByVal Prezzo As Variant, _
                                    Optional ByVal NumColli As Variant, Optional ByVal Quantita As Variant, Optional ByVal ScontoImporto As Variant, _
                                    Optional ByVal PercentualeSconto1 As Variant, Optional ByVal PercentualeSconto2 As Variant, _
                                    Optional ByVal PercentualeSconto3 As Variant, Optional ByVal PercentualeSconto4 As Variant, _
                                    Optional ByVal PercentualeSconto5 As Variant, Optional ByVal PercentualeSconto6 As Variant, _
                                    Optional ByVal CodiceTestoFisso As Variant, Optional ByVal DesRiga As Variant, _
                                    Optional ByVal PercentualeProvvigioni As Variant, Optional ByVal CodiceLotto As Variant, _
                                    Optional ByVal ScadenzaLotto As Variant, Optional ByVal CodiceUbicazione As Variant, _
                                    Optional ByVal RiferimentoNumreg As Variant, Optional ByVal RiferimentoProgriga As Variant, _
                                    Optional ByVal RiferimentoTipoDoc As Variant, Optional ByVal RiferimentoSottoTipoDoc As Variant, _
                                    Optional ByVal RiferimentoNumDoc As Variant, Optional ByVal RiferimentoSezDoc As Variant, _
                                    Optional ByVal RiferimentoDocBis As Variant, Optional ByVal RiferimentoDataDoc As Variant, _
                                    Optional ByVal RiferimentoProgrigaMaga As Variant, Optional ByVal RiferimentoProgrigaLotto As Variant, _
                                    Optional ByVal VDS As Variant, Optional ByVal CDC As Variant, Optional ByVal Commessa As Variant, _
                                    Optional ByVal UM1 As Variant, Optional ByVal UM2 As Variant, Optional ByVal FatConv As Variant, _
                                    Optional ByVal PZConf As Variant, Optional ByVal PesoL As Variant, Optional ByVal PesoN As Variant, _
                                    Optional ByVal Tara As Variant, Optional ByVal SottoCommessa As Variant, Optional ByVal DataConsegna As Variant, Optional ByVal StatoRigaOfferta As Variant)



    On Error GoTo Err_AddNewRigaDocumento

    Dim NuovoImportoProvvigioni As Single
    Errore = 0
    ErrDescr = ""

    With ClsRegDoc
        
      ' If Not Pbol_EseguitaAddNew Then
       ' problema trinchero
      .AddNewDocCorpo False, , , True
       'End If
       'Pbol_EseguitaAddNew = False
      
      If Not IsMissing(CodiceAliquotaIVA) Then
        .CambiaDocCorpoAliva CodiceAliquotaIVA
      End If
      
      Select Case TipoRiga
         Case tsImportDocumentiArticoloMagazzino
             .rstDocCorpo("DO30_INDTIPORIGA").value = 0
         Case tsImportDocumentiArticoloManuale
             .rstDocCorpo("DO30_INDTIPORIGA").value = 1
         Case tsImportDocumentiFaseLavorazione
             .rstDocCorpo("DO30_INDTIPORIGA").value = 10
         Case tsImportDocumentiRigaDescrittiva
             .rstDocCorpo("DO30_INDTIPORIGA").value = 2
         Case tsImportDocumentiSpesa
             .rstDocCorpo("DO30_INDTIPORIGA").value = 4
         Case tsImportDocumentiTestoFisso
             .rstDocCorpo("DO30_INDTIPORIGA").value = 6
      End Select
      
      If Not IsMissing(UM1) Then
        .rstDocCorpo("DO30_UM1").value = UM1
      End If
      
      If Not IsMissing(UM2) Then
        .rstDocCorpo("DO30_UM2").value = UM2
      End If

      If Not IsMissing(FatConv) Then
        .rstDocCorpo("DO30_FATCONV").value = FatConv
      End If

      If Not IsMissing(PZConf) Then
        .rstDocCorpo("DO30_PZCONF").value = PZConf
      End If

      If Not IsMissing(PesoL) Then
        .rstDocCorpo("DO30_PESOL").value = PesoL
      End If
      
      If Not IsMissing(PesoN) Then
        .rstDocCorpo("DO30_PESON").value = PesoN
      End If

      If Not IsMissing(Tara) Then
        .rstDocCorpo("DO30_TARA").value = Tara
      End If

'      If Not IsMissing(Capacita) Then
'        .rstDocCorpo("DO30_CAPACITA").Value = Capacita
'      End If

     
      
      If Not IsMissing(CodiceArticolo) Then
        .CambiaDocCorpoCodart CodiceArticolo, CodiceVarianteArticolo, True
      End If
      
       If Not IsMissing(DesRiga) Then
        '.rstDocCorpo("DO30_DESCART").Value = DesRiga
        .CambiaDocCorpoDescrizioniArticolo DesRiga, Null
        
      End If
      
      If Not IsMissing(CodiceDeposito) Then
         .CambiaDocCorpoCodDep CodiceDeposito
      End If
      
      If Not IsMissing(CodiceSpesa) Then
         .CambiaDocCorpoCodSpesa CodiceSpesa
      End If
      
      If Not IsMissing(CodiceFaseLavorazione) Then
         .CambiaDocCorpoFase CodiceFaseLavorazione
      End If
      
      If Not IsMissing(NumColli) Then
         .CambiaDocCorpoColli NumColli
      End If
      
      If Not IsMissing(Quantita) Then
         .CambiaDocCorpoQta1 Quantita, False, False, False
      End If
      
      If Not IsMissing(MaggiorazioneImporto) Then
         .CambiaDocCorpoMagImp MaggiorazioneImporto
      End If
      
      If Not IsMissing(PercentualeMaggiorazione1) Then
         .CambiaDocCorpoMagPer1 PercentualeMaggiorazione1
      End If
      
      If Not IsMissing(PercentualeMaggiorazione2) Then
         .CambiaDocCorpoMagPer2 PercentualeMaggiorazione2
      End If
      
      If Not IsMissing(Prezzo) Then
         .CambiaDocCorpoPrezzo1 Prezzo
      End If
      
      If Not IsMissing(ScontoImporto) Then
         .CambiaDocCorpoScImp ScontoImporto
      End If
      
      If Not IsMissing(PercentualeSconto1) Then
         .CambiaDocCorpoScPer1 PercentualeSconto1
      End If
      
      If Not IsMissing(PercentualeSconto2) Then
         .CambiaDocCorpoScPer2 PercentualeSconto2
      End If
      
      If Not IsMissing(PercentualeSconto3) Then
         .CambiaDocCorpoScPer3 PercentualeSconto3
      End If
      
      If Not IsMissing(PercentualeSconto4) Then
         .CambiaDocCorpoScPer4 PercentualeSconto4
      End If
      
      If Not IsMissing(PercentualeSconto5) Then
         .CambiaDocCorpoScPer5 PercentualeSconto5
      End If
      
      If Not IsMissing(PercentualeSconto6) Then
         .CambiaDocCorpoScPer6 PercentualeSconto6
      End If
      
      If Not IsMissing(CodiceTestoFisso) Then
         .CambiaDocCorpoTestoFix CodiceTestoFisso
      End If
      
'      If Not IsMissing(DataConsegna) Then
'         .Cambiadoc CodiceTestoFisso
'      End If
      
      
      .UpdateDocCorpo , True

      'COAN
      If Not IsMissing(VDS) Or Not IsMissing(CDC) Then
      '          If NVL(VDS, "") <> "" Or NVL(CDC, "") <> "" Then
        If .rstDocCorpoCoin.RecordCount = 0 Then
            Call .AddNewDocCorpoCoin
        End If

        If Not IsMissing(CDC) Then
          .rstDocCorpoCoin.Fields("DO53_PDCDEST_PC01").value = ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsParCOIN.CodicePDC
          .rstDocCorpoCoin.Fields("DO53_CONTODEST_PC03").value = CDC
        End If

        If Not IsMissing(VDS) Then
          .rstDocCorpoCoin.Fields("DO53_PDCVDSDEST_PC01").value = ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsParCOIN.CodiceVdS
          .rstDocCorpoCoin.Fields("DO53_VDSDEST_PC03").value = VDS
        End If

        .UpdateDocCorpoCoin

      End If


      If Not IsMissing(Commessa) Then
      If .rstDocCorpoProge.RecordCount = 0 Then
          Call .AddNewDocCorpoProge
      End If
  
      .rstDocCorpoProge.Fields("DO41_CODCOMMESSA").value = Commessa
      If Not IsMissing(SottoCommessa) Then
        
        .rstDocCorpoProge.Fields("DO41_CODSCOMMESSA").value = SottoCommessa
      End If
      .UpdateDocCorpoProge
    End If
    
    If Not IsMissing(DataConsegna) Then
    .rstDocCorpoOrd.Fields("DO31_DATACONS").value = CDate(DataConsegna)
    End If
  
      
'      If Not IsMissing(CodiceLotto) Then
'        If .rstDocCorpoLot.RecordCount = 0 Then
'          Call .AddNewDocCorpoLot
'        End If
'         .rstDocCorpoLot.Fields("DO52_CODLOTTO_MG4G").Value = CodiceLotto
'         .rstDocCorpoLot.Fields("DO52_SCADENZA").Value = ScadenzaLotto
'
'         .UpdateDocCorpoLot
'      End If

       If Not IsMissing(StatoRigaOfferta) Then
        If .rstDocCorpoPers.RecordCount = 0 Then
          Call .AddNewDocCorpoPers
        End If
       
           .rstDocCorpoPers.Fields("DO35_DITTA_CG18").value = CodiceDitta
           .rstDocCorpoPers.Fields("DO35_NUMREG_CO99").value = .rstDocCorpo("DO30_NUMREG_CO99")
           .rstDocCorpoPers.Fields("DO35_PROGRIGA").value = .rstDocCorpo("DO30_PROGRIGA")
           .rstDocCorpoPers.Fields("DO35_ALFPERS1").value = StatoRigaOfferta
       .UpdateDocCorpoPers
       End If
       
      If Not IsMissing(RiferimentoNumreg) And Not IsMissing(RiferimentoProgriga) Then
        If NVL(RiferimentoNumreg, "") <> "" And NVL(RiferimentoProgriga, "") <> "" Then
          'Inserimento riferimenti
          .AddNewDocCorpoRif False

          .rstDocCorpoRif.Fields("DO33_NUMREG_CO99").value = .rstDocCorpo("DO30_NUMREG_CO99")
          .rstDocCorpoRif.Fields("DO33_PROGRIGA").value = .rstDocCorpo("DO30_PROGRIGA")
          .rstDocCorpoRif.Fields("DO33_PROGRIF").value = 1

          If Not IsMissing(RiferimentoNumreg) Then
            .rstDocCorpoRif.Fields("DO33_NUMREGRIF_CO99").value = Trim(RiferimentoNumreg)
          End If

          If Not IsMissing(RiferimentoProgriga) Then
            .rstDocCorpoRif.Fields("DO33_PROGRIGARIF_DO30").value = RiferimentoProgriga
          End If

          If Not IsMissing(RiferimentoTipoDoc) Then
            .rstDocCorpoRif.Fields("DO33_INDTIPODOC").value = RiferimentoTipoDoc
          End If

          If Not IsMissing(RiferimentoSottoTipoDoc) Then
            .rstDocCorpoRif.Fields("DO33_INDSTIPODOC").value = RiferimentoSottoTipoDoc
          End If

          If Not IsMissing(RiferimentoNumDoc) Then
            .rstDocCorpoRif.Fields("DO33_NUMNSDOC").value = RiferimentoNumDoc
          End If

          If Not IsMissing(RiferimentoSezDoc) Then
            .rstDocCorpoRif.Fields("DO33_SEZNSDOC").value = RiferimentoSezDoc
          End If

          If Not IsMissing(RiferimentoDocBis) Then
            .rstDocCorpoRif.Fields("DO33_FLGNSDOCBIS").value = RiferimentoDocBis
          End If

          If Not IsMissing(RiferimentoDataDoc) Then
            .rstDocCorpoRif.Fields("DO33_DATANSDOC").value = RiferimentoDataDoc
          End If

          .rstDocCorpoRif.Fields("DO33_INDQTAORDRES").value = 0
          .rstDocCorpoRif.Fields("DO33_FLGAGGQTATRASF").value = 1
          .rstDocCorpoRif.Fields("DO33_FLGSTORNOQTAVAL").value = 1

          If Not IsMissing(Quantita) Then
            .rstDocCorpoRif.Fields("DO33_QTA1MOV").value = NVL(.rstDocCorpo("DO30_QTA1").value, 0)
            .rstDocCorpoRif.Fields("DO33_QTA2MOV").value = NVL(.rstDocCorpo("DO30_QTA2").value, 0)
'             .CambiaDocCorpoQta1 Quantita
          End If
          
          If Not IsMissing(NumColli) Then
             .rstDocCorpoRif.Fields("DO33_COLLIMOV").value = NumColli
          End If
 
          .UpdateDocCorpoRif

        End If
      End If
      
'      .RicalcolaCostoRiga
'      .RicalcolaPrezzo
'      .RicalcolaImportoRiga
'      .RicalcolaTotaliDaRigheCorpo
      
      .UpdateDocCorpo , True
            
      NumregDocumento = .rstDocCorpo("DO30_NUMREG_CO99")
      ProgRigaDocumento = .rstDocCorpo("DO30_PROGRIGA")
      
      
      
    End With

   Errore = ClsRegDoc.Stato
   ErrDescr = ClsRegDoc.StatoDesc

Exit Function
Err_AddNewRigaDocumento:
   Errore = tsErrProgram
   ErrDescr = Err.Description
End Function

'
' Aggiunta Riga del corpo documento
'
Public Function AddNewRigaLotto(Optional ByVal CodiceArticolo As Variant, Optional ByVal CodiceVarianteArticolo As Variant, Optional ByVal CodiceDeposito As Variant, _
                                Optional ByVal CodiceLotto As Variant, Optional ByVal DescrizioneLotto As Variant, Optional ByVal ScadenzaLotto As Variant, _
                                Optional ByVal CodiceUbicazione As Variant, Optional ByVal Quantita As Variant, _
                                Optional ByVal NumColli As Variant, Optional ByVal RiferimentoProgrigaLotto As Variant, _
                                Optional ByVal ElemTracciabilita1 As Variant, Optional ByVal FattConv As Variant)

    On Error GoTo Err_AddNewRigaLotto

    Errore = 0
    ErrDescr = ""
    
    With ClsRegDoc
      If Not IsMissing(CodiceLotto) Or Not IsMissing(CodiceUbicazione) Then
        'Inserimento Lotti/Ubicazioni
        .AddNewDocCorpoLot False
        
        .rstDocCorpoLot.Fields("DO52_DITTA_CG18").value = CodiceDitta
'        .rstDocCorpoLot.Fields("DO52_NUMREG_CO99").Value = .rstDocCorpo("DO30_NUMREG_CO99")
'        .rstDocCorpoLot.Fields("DO52_PROGRIGA").Value = .rstDocCorpo("DO30_PROGRIGA")
        .rstDocCorpoLot.Fields("DO52_NUMREG_CO99").value = NumregDocumento
        .rstDocCorpoLot.Fields("DO52_PROGRIGA").value = ProgRigaDocumento
        .rstDocCorpoLot.Fields("DO52_PROG_MG4F").value = 1
        .rstDocCorpoLot.Fields("DO52_PROG").value = RiferimentoProgrigaLotto
        

''''''SELECT     DO52_DITTA_CG18, DO52_NUMREG_CO99, DO52_PROGRIGA, DO52_PROG_MG4F, DO52_PROG, DO52_CODART_MG66, DO52_OPZIONE_MG5E,
''''''                      DO52_CODDEP_MG58, DO52_CODLOTTO_MG4G, DO52_SCADENZA, DO52_SERNUM, DO52_CODPALLETS, DO52_CODPROG_PD14, DO52_SOTPROG_PD14,
''''''                      DO52_UBICAZ_MG97, DO52_CODCONFEZ_MG96, DO52_CODBAGNOMAT, DO52_QTA1, DO52_QTA2, DO52_QTA1CONS, DO52_QTA2CONS, DO52_TIPOCF_CG44,
''''''                      DO52_CLIFOR_CG44, DO52_NOTE, DO52_CODARTPF_MG66, DO52_OPZIONEPF_MG5E, DO52_CODLOTTOPF_MG4G, DO52_NUMREGRIF_CO99,
''''''                      DO52_PROGRIGARIF_DO30 , DO52_PROGRIF_DO52, DO52_IDMEDIA_CG99, DO52_CODSSCC, DO52_CAUSMAG_MG51, DO52_FLGNONPIUEV
''''''From
'''''''    DO52_PROG_MG4F, DO52_PROG,
''''''                            Numreg, _
''''''                            ProgRiga, _
''''''
''''''
                                
                                
'        .rstDocCorpoLot.Fields("DO52_NUMREG_CO99").Value = NumregDocumento
'        .rstDocCorpoLot.Fields("DO52_PROGRIGA").Value = ProgRigaDocumento
'        .rstDocCorpoLot.Fields("DO52_PROG_MG4F").Value = 1
'        .rstDocCorpoLot.Fields("DO52_PROG").Value = 1

        If Not IsMissing(CodiceArticolo) Then
          .rstDocCorpoLot.Fields("DO52_CODART_MG66").value = CodiceArticolo
        End If
      
        If Not IsMissing(CodiceVarianteArticolo) Then
          .rstDocCorpoLot.Fields("DO52_OPZIONE_MG5E").value = CodiceVarianteArticolo
        Else
          .rstDocCorpoLot.Fields("DO52_OPZIONE_MG5E").value = ""
        End If
      
        If Not IsMissing(CodiceDeposito) Then
          .rstDocCorpoLot.Fields("DO52_CODDEP_MG58").value = CodiceDeposito
        End If
        
        If Not IsMissing(CodiceLotto) Then
          If NVL(CodiceLotto, "") <> "" Then
            .rstDocCorpoLot.Fields("DO52_CODLOTTO_MG4G").value = CodiceLotto
            Call CreaRecordMG4G_ANAGRLOTTI(CodiceDitta, CodiceArticolo, CodiceVarianteArticolo, CodiceLotto, NVL(DescrizioneLotto, ""))
          End If
        End If
        
        If Not IsMissing(ElemTracciabilita1) Then
          If NVL(ElemTracciabilita1, "") <> "" Then
            Call CreaRecordMG4L_ELEMTRALIM(CodiceDitta, ElemTracciabilita1, "")
            Call CreaRecordMG4P_LOTTRALIM(CodiceDitta, CodiceArticolo, CodiceVarianteArticolo, CodiceLotto, NVL(DescrizioneLotto, ""), ElemTracciabilita1)
          End If
        End If

        If Not IsMissing(ScadenzaLotto) Then
          If NVL(ScadenzaLotto, "") <> "" Then
            .rstDocCorpoLot.Fields("DO52_SCADENZA").value = ScadenzaLotto
          End If
        End If

        If Not IsMissing(CodiceUbicazione) Then
          If NVL(CodiceUbicazione, "") <> "" Then
            .rstDocCorpoLot.Fields("DO52_UBICAZ_MG97").value = CodiceUbicazione
          End If
        End If
        
        If Not IsMissing(Quantita) Then
           .rstDocCorpoLot.Fields("DO52_QTA1").value = Quantita
        End If
        
        If Not IsMissing(Quantita) And Not IsMissing(FattConv) Then
           .rstDocCorpoLot.Fields("DO52_QTA2").value = Quantita * FattConv
        End If
        
        .rstDocCorpoLot.Fields("DO52_FLGNONPIUEV").value = 0
        
        
        .UpdateDocCorpoLot
        
        .UpdateDocCorpo , True
        
      End If
           
           
    End With

   Errore = ClsRegDoc.Stato
   ErrDescr = ClsRegDoc.StatoDesc

Exit Function
Err_AddNewRigaLotto:
   Errore = tsErrProgram
   ErrDescr = Err.Description
End Function


'
' Aggiorna Riga del corpo lotto documento
'
Public Function UpdateRigaLotto(ByVal rNumReg As String, ByVal rProgRiga As String, _
                                ByVal rNumProgMG4F As String, ByVal rProgRigaLotto As String, _
                                Optional ByVal CodiceArticolo As Variant, Optional ByVal CodiceVarianteArticolo As Variant, Optional ByVal CodiceDeposito As Variant, _
                                Optional ByVal CausaleMagazzino As Variant, Optional ByVal UM1 As Variant, _
                                Optional ByVal CodiceLotto As Variant, Optional ByVal DescrizioneLotto As Variant, Optional ByVal ScadenzaLotto As Variant, _
                                Optional ByVal CodiceUbicazione As Variant, Optional ByVal Quantita As Variant, _
                                Optional ByVal NumColli As Variant, Optional ByVal RiferimentoProgrigaLotto As Variant, _
                                Optional ByVal ElemTracciabilita1 As Variant)

    On Error GoTo Err_AddNewRigaLotto

    Dim strFiltro As String
    
    strFiltro = "DO52_NUMREG_CO99 = '" & rNumReg & "'"
    strFiltro = strFiltro & " AND DO52_PROGRIGA = " & rProgRiga
    strFiltro = strFiltro & " AND DO52_PROG_MG4F = " & rNumProgMG4F
    strFiltro = strFiltro & " AND DO52_PROG = " & rProgRigaLotto
    
    Errore = 0
    ErrDescr = ""
    
    With ClsRegDoc
      If Not IsMissing(CodiceLotto) Or Not IsMissing(CodiceUbicazione) Then
        'Inserimento Lotti/Ubicazioni
        .OpenDocCorpoLot strFiltro
        
        If .rstDocCorpoLot.RecordCount = 0 Then
            .AddNewDocCorpoLot False
        End If
        
'        .rstDocCorpoLot.Fields("DO52_NUMREG_CO99").Value = rNumReg
'        .rstDocCorpoLot.Fields("DO52_PROGRIGA").Value = rProgRiga
'        .rstDocCorpoLot.Fields("DO52_PROG_MG4F").Value = rNumProgMG4F
'        .rstDocCorpoLot.Fields("DO52_PROG").Value = rProgRigaLotto
        
        If Not IsMissing(CodiceArticolo) Then
          .rstDocCorpoLot.Fields("DO52_CODART_MG66").value = CodiceArticolo
        End If
      
        If Not IsMissing(CodiceVarianteArticolo) Then
          .rstDocCorpoLot.Fields("DO52_OPZIONE_MG5E").value = CodiceVarianteArticolo
        Else
          .rstDocCorpoLot.Fields("DO52_OPZIONE_MG5E").value = ""
        End If
      
        If Not IsMissing(CodiceDeposito) Then
          .rstDocCorpoLot.Fields("DO52_CODDEP_MG58").value = CodiceDeposito
        End If
        
        If Not IsMissing(CodiceLotto) Then
          If NVL(CodiceLotto, "") <> "" Then
            .rstDocCorpoLot.Fields("DO52_CODLOTTO_MG4G").value = CodiceLotto
            Call CreaRecordMG4G_ANAGRLOTTI(CodiceDitta, CodiceArticolo, CodiceVarianteArticolo, CodiceLotto, NVL(DescrizioneLotto, ""))
          End If
        End If
        
        If Not IsMissing(ElemTracciabilita1) Then
          If NVL(ElemTracciabilita1, "") <> "" Then
            Call CreaRecordMG4L_ELEMTRALIM(CodiceDitta, ElemTracciabilita1, "")
            Call CreaRecordMG4P_LOTTRALIM(CodiceDitta, CodiceArticolo, CodiceVarianteArticolo, CodiceLotto, NVL(DescrizioneLotto, ""), ElemTracciabilita1)
          End If
        End If
        
        If Not IsMissing(ScadenzaLotto) Then
          If NVL(ScadenzaLotto, "") <> "" Then
            .rstDocCorpoLot.Fields("DO52_SCADENZA").value = ScadenzaLotto
          End If
        End If
      
        If Not IsMissing(CodiceUbicazione) Then
          If NVL(CodiceUbicazione, "") <> "" Then
            .rstDocCorpoLot.Fields("DO52_UBICAZ_MG97").value = CodiceUbicazione
          End If
        End If
                
        If Not IsMissing(Quantita) Then
           .rstDocCorpoLot.Fields("DO52_QTA1").value = Quantita
        End If
        
        If Not IsMissing(CausaleMagazzino) Then
           .rstDocCorpoLot.Fields("DO52_CAUSMAG_MG51").value = CausaleMagazzino
        End If
        
        .UpdateDocCorpoLot
        
        .RicalcolaCostoRiga
        .RicalcolaPrezzo
        .RicalcolaImportoRiga
        .RicalcolaTotaliDaRigheCorpo , , , False
        
        .UpdateDocCorpo , True, , False
        
'        .UpdateDocCorpo , True
'        .RicalcolaCostoRiga
'        .RicalcolaPrezzo
'        .RicalcolaImportoRiga
'        .RicalcolaTotaliDaRigheCorpo , , , False
'
'        .UpdateDocCorpo , True
        
      End If
      
    End With

   Errore = ClsRegDoc.Stato
   ErrDescr = ClsRegDoc.StatoDesc

Exit Function
Err_AddNewRigaLotto:
   Errore = tsErrProgram
   ErrDescr = Err.Description
End Function


Private Sub CreaRecordMG4G_ANAGRLOTTI(fcodiceditta, fcodicearticolo, fvariantearticolo, fcodicelotto, fdescrizionelotto)
    Dim strSQL As String
On Error GoTo ErrTrap
    strSQL = "IF NOT EXISTS(SELECT TOP 1 1 FROM MG4G_ANAGRLOTTI "
    strSQL = strSQL & " WHERE MG4G_DITTA_CG18 = " & fcodiceditta
    strSQL = strSQL & " AND MG4G_CODART_MG66 = '" & fcodicearticolo & "' "
    If Not IsMissing(fvariantearticolo) Then
      strSQL = strSQL & " AND MG4G_OPZIONE_MG5E = '" & fvariantearticolo & "' "
    Else
      strSQL = strSQL & " AND MG4G_OPZIONE_MG5E = '' "
    End If
    strSQL = strSQL & " AND MG4G_CODLOTTO = '" & Trim(fcodicelotto) & "') "
    strSQL = strSQL & " INSERT MG4G_ANAGRLOTTI "
    strSQL = strSQL & " (MG4G_DITTA_CG18, MG4G_CODART_MG66, "
    strSQL = strSQL & "  MG4G_OPZIONE_MG5E, "
    strSQL = strSQL & "  MG4G_CODLOTTO"
    
    If Not IsMissing(fdescrizionelotto) Then
      strSQL = strSQL & "  , MG4G_DESCLOTTO "
    End If
    
    strSQL = strSQL & ") VALUES ("
    strSQL = strSQL & fcodiceditta & ", '" & fcodicearticolo & "', "
    
    If Not IsMissing(fvariantearticolo) Then
      strSQL = strSQL & "'" & fvariantearticolo & "',"
    Else
      strSQL = strSQL & "'',"
    End If
    
    strSQL = strSQL & "'" & Trim(fcodicelotto) & "'"
    
    If Not IsMissing(fdescrizionelotto) Then
      strSQL = strSQL & ", '" & Trim(fdescrizionelotto) & "'"
    End If
    
    strSQL = strSQL & ")"
    
    Gcon_Connect.Execute strSQL, , adCmdText
Exit Sub
ErrTrap:
Debug.Print Err.Description

'    gclsTracciato.AppendToLog "Errore", Err.Description & " in CreaRecordMG4G_ANAGRLOTTI." & vbCrLf & "Articolo: " & fcodicearticolo & vbCrLf & "Variante: " & fvariantearticolo & vbCrLf & "Codice Lotto: " & fcodicelotto & vbCrLf & "Descrizione: " & fdescrizionelotto & vbCrLf & "in riga " & glngProgressivoRiga & vbCrLf & _
'        "Ditta: " & fcodiceditta & vbCrLf & _
'        "Codice articolo: " & fcodicearticolo & vbCrLf & _
'        "Codice Variante: " & fvariantearticolo & vbCrLf & _
'        "Codice Lotto: " & fcodicelotto & vbCrLf & _
'        "Descrizione lotto: " & fdescrizionelotto, False
End Sub

Private Sub CreaRecordMG4L_ELEMTRALIM(fcodiceditta, fcodiceelem1, fdescrizioneelem1)
On Error GoTo ErrTrap
    
    Dim strSQL As String

    strSQL = "IF NOT EXISTS(SELECT TOP 1 1 FROM MG4L_ELEMTRALIM "
    strSQL = strSQL & "      WHERE MG4L_DITTA_CG18 = " & fcodiceditta
    strSQL = strSQL & "        AND MG4L_CODELEM_MG4I = 1 "
    strSQL = strSQL & "        AND MG4L_CODICE = '" & fcodiceelem1 & "') "
    strSQL = strSQL & "     INSERT MG4L_ELEMTRALIM "
    strSQL = strSQL & "     ("
    strSQL = strSQL & "       MG4L_DITTA_CG18, MG4L_CODELEM_MG4I, MG4L_CODICE, MG4L_DESCR "
    strSQL = strSQL & "     ) VALUES ("
    strSQL = strSQL & "       " & fcodiceditta & ", 1, '" & fcodiceelem1 & "', '" & fdescrizioneelem1 & "'"
    strSQL = strSQL & "     ) "
    Gcon_Connect.Execute strSQL, , adCmdText
    
    
Exit Sub
ErrTrap:
Debug.Print Err.Description
'    gclsTracciato.AppendToLog "Errore", Err.Description & " in CreaRecordMG4G_ANAGRLOTTI." & vbCrLf & "Articolo: " & fcodicearticolo & vbCrLf & "Variante: " & fvariantearticolo & vbCrLf & "Codice Lotto: " & fcodicelotto & vbCrLf & "Descrizione: " & fdescrizionelotto & vbCrLf & "in riga " & glngProgressivoRiga & vbCrLf & _
'        "Ditta: " & fcodiceditta & vbCrLf & _
'        "Codice articolo: " & fcodicearticolo & vbCrLf & _
'        "Codice Variante: " & fvariantearticolo & vbCrLf & _
'        "Codice Lotto: " & fcodicelotto & vbCrLf & _
'        "Descrizione lotto: " & fdescrizionelotto, False
End Sub

Private Sub CreaRecordMG4P_LOTTRALIM(fcodiceditta, fcodicearticolo, fvariantearticolo, fcodicelotto, fdescrizionelotto, fcodiceelem1)
    Dim strSQL As String
On Error GoTo ErrTrap
    
    strSQL = "IF NOT EXISTS(SELECT TOP 1 1 FROM MG4P_LOTTRALIM "
    strSQL = strSQL & "      WHERE MG4P_DITTA_CG18 = " & fcodiceditta
    strSQL = strSQL & "        AND MG4P_CODART_MG66 = '" & fcodicearticolo & "'"
    
    If Not IsMissing(fvariantearticolo) Then
      strSQL = strSQL & " AND MG4P_OPZIONE_MG5E = '" & fvariantearticolo & "' "
    Else
      strSQL = strSQL & " AND MG4P_OPZIONE_MG5E = '' "
    End If
    
    strSQL = strSQL & "        AND MG4P_CODLOTTO_MG4G = '" & fcodicelotto & "'"
    strSQL = strSQL & "     ) "
    strSQL = strSQL & "     INSERT MG4P_LOTTRALIM "
    strSQL = strSQL & "     ("
    strSQL = strSQL & "       MG4P_DITTA_CG18, MG4P_CODART_MG66, "
    strSQL = strSQL & "       MG4P_OPZIONE_MG5E , "
    strSQL = strSQL & "       MG4P_CODLOTTO_MG4G , MG4P_CODELEM_MG4I, MG4P_CODICE_MG4L "
    
    strSQL = strSQL & "     ) VALUES ("
    strSQL = strSQL & "       " & fcodiceditta & ", '" & fcodicearticolo & "', "
    
    If Not IsMissing(fvariantearticolo) Then
      strSQL = strSQL & "     '" & fvariantearticolo & "',"
    Else
      strSQL = strSQL & "     '',"
    End If
    
    strSQL = strSQL & "      '" & fcodicelotto & "', 1, CAST ( '" & fcodiceelem1 & "' AS CHAR(20))"
    strSQL = strSQL & "     ) "
    Gcon_Connect.Execute strSQL, , adCmdText
     
    
Exit Sub
ErrTrap:
Debug.Print Err.Description
'    gclsTracciato.AppendToLog "Errore", Err.Description & " in CreaRecordMG4G_ANAGRLOTTI." & vbCrLf & "Articolo: " & fcodicearticolo & vbCrLf & "Variante: " & fvariantearticolo & vbCrLf & "Codice Lotto: " & fcodicelotto & vbCrLf & "Descrizione: " & fdescrizionelotto & vbCrLf & "in riga " & glngProgressivoRiga & vbCrLf & _
'        "Ditta: " & fcodiceditta & vbCrLf & _
'        "Codice articolo: " & fcodicearticolo & vbCrLf & _
'        "Codice Variante: " & fvariantearticolo & vbCrLf & _
'        "Codice Lotto: " & fcodicelotto & vbCrLf & _
'        "Descrizione lotto: " & fdescrizionelotto, False
End Sub
'
'
'
'Public Sub AggiornaStatoEvasione(ByVal rNumReg As String)
'
'On Error GoTo ErrTrap
'
'  With ClsRegDoc
'    .AggiornaStatoEvasioneDocumento (rNumReg)
'
''    .UpdateDocTestata
'  End With
'
''  If ClsRegDoc.rstDocTestata.RecordCount > 0 Then
''      ClsRegDoc.DeleteDocTestata
''      If ClsRegDoc.Stato <> 0 Then
''        ErrDescr = ClsRegDoc.StatoDesc & vbNewLine & _
''                 "Eliminazione doc. " & rCodDoc & " n.r." & rNumReg
''        Errore = tsErrDocEliminazione
''      End If
''  End If
'
'  Exit Sub
'ErrTrap:
'    Errore = tsErrProgram
'    ErrDescr = Err.Description
'End Sub

Private Function NVL(Valore As Variant, ValIfNull As Variant) As Variant
  ' utility function
  On Error Resume Next
  
  If IsEmpty(Valore) Or IsNull(Valore) Then
    NVL = ValIfNull
  Else
    If Trim(CStr(Valore)) = "" Then
        NVL = ValIfNull
    Else
        NVL = Valore
    End If
  End If
  
  Err.Clear
End Function

Private Sub Class_Terminate()
  On Error Resume Next
  
  Set Gcon_Connect = Nothing
  Set ActiveInterface = Nothing
  
  Err.Clear
End Sub

Private Function GeneraDocumentoCollegato(ByVal CodiceDocumentoCollegato As String) As String
    '
    ' Variabili locali
    '
    Dim ClsRegDoc_Col                As MGBO_REGDOCUMENTI.CLSMG_REGDOC
    Dim rst_DocCorpo_Clone          As ADODB.Recordset
    Dim rst_DocCorpoOrd_Clone       As ADODB.Recordset
    Dim rst_DocCorpoProge_Clone     As ADODB.Recordset
    Dim rst_DocCorpoLot_Clone       As ADODB.Recordset
    Dim ADOField                    As ADODB.Field
    Dim Pint_Index                  As Integer
    Dim Pstr_Sql                    As String
    Dim Pstr_Filter                 As String
    Dim Prst_PersLotti              As ADODB.Recordset
    Dim Pbol_MovLotti               As Boolean
    Dim Pbol_DO52_Div               As Boolean
    Dim Pstr_Filter_MG4F            As String
    Dim Pint_Num_Filter_MG4F        As Integer
    Dim Parr_Filter_MG4F            As Variant
    Dim Plng_NumRec52_Orig          As Long
    Dim Plng_NumRec52_Filter        As Long
    Dim Pbol_SoloProgColl           As Boolean
    Dim Prst_SezDefault             As ADODB.Recordset
    Dim Pstr_SezDefault             As String
    Dim Pvar_Progetto               As Variant
    '
    ' Verifico se movimentare i lotti sul documento collegato
    '
    If ClsRegDoc.RegDocIn.GestDocCorpoLot Then
        Pstr_Sql = " SELECT" & _
                   "    MG4H_PROG_MG4F" & _
                   " FROM" & _
                   "    MG4H_PERSDOCLOTTI WITH (NOLOCK)" & _
                   " WHERE" & _
                   "    MG4H_DITTA_CG18 = " & CodiceDitta & " AND" & _
                   "    MG4H_CODDOCUM_MG3G = '" & RTrimN(CodiceDocumentoCollegato) & "' AND" & _
                   "    MG4H_INDTIPOGES > 0"
        Set Prst_PersLotti = Gcon_Connect.Execute(Pstr_Sql, , adCmdText)
        Pbol_MovLotti = Not Prst_PersLotti.EOF
        If Pbol_MovLotti Then
            Pbol_DO52_Div = False
            Pint_Num_Filter_MG4F = 0
            Pstr_Filter_MG4F = ""
            While Not Prst_PersLotti.EOF
                Pint_Num_Filter_MG4F = Pint_Num_Filter_MG4F + 1
                If Pint_Num_Filter_MG4F = 1 Then
                    ReDim Parr_Filter_MG4F(1 To 1)
                Else
                    ReDim Preserve Parr_Filter_MG4F(1 To Pint_Num_Filter_MG4F)
                End If
                If Len(Pstr_Filter_MG4F) > 0 Then
                    Pstr_Filter_MG4F = Pstr_Filter_MG4F & " OR "
                End If
                Pstr_Filter_MG4F = Pstr_Filter_MG4F & "DO52_PROG_MG4F = " & CDecN(Prst_PersLotti("MG4H_PROG_MG4F").value)
                Parr_Filter_MG4F(Pint_Num_Filter_MG4F) = CStr(CDecN(Prst_PersLotti("MG4H_PROG_MG4F").value))
                Prst_PersLotti.MoveNext
            Wend
        Else
            Pbol_DO52_Div = True
        End If
        Prst_PersLotti.Close
        Set Prst_PersLotti = Nothing
    Else
        Pbol_MovLotti = False
        Pbol_DO52_Div = False
    End If
    '
    ' Azzero il booleano che indica la presenza del solo progetto collegato
    '
    Pbol_SoloProgColl = False
    '
    ' Istanzio il Bo per la memorizzazione del documento collegato
    '
    Set ClsRegDoc_Col = New MGBO_REGDOCUMENTI.CLSMG_REGDOC
    ClsRegDoc_Col.RegDocIn.Transaction = False
    Set ClsRegDoc_Col.RegDocIn.ActiveInterface = ActiveInterface
    Set ClsRegDoc_Col.RegDocIn.Connection = Gcon_Connect
    ClsRegDoc_Col.RegDocIn.Ditta = CodiceDitta
    ClsRegDoc_Col.RegDocIn.Documento = CodiceDocumentoCollegato
    ClsRegDoc_Col.RegDocIn.GestDocTestata = True
    ClsRegDoc_Col.RegDocIn.GestDocDatiAcc = True
    ClsRegDoc_Col.RegDocIn.GestDocTotali = True
    ClsRegDoc_Col.RegDocIn.GestDocTestaProge = ClsRegDoc.RegDocIn.GestDocTestaProge
    ClsRegDoc_Col.RegDocIn.GestDocCorpo = True
    ClsRegDoc_Col.RegDocIn.GestDocCorpoOrd = True
    ClsRegDoc_Col.RegDocIn.GestDocCorpoProge = ClsRegDoc.RegDocIn.GestDocCorpoProge
    ClsRegDoc_Col.RegDocIn.GestDocCorpoLot = Pbol_MovLotti
    ClsRegDoc_Col.Initialize
    ClsRegDoc_Col.OpenDocTestata "1=0"
    ClsRegDoc_Col.OpenDocDatiAcc "1=0"
    ClsRegDoc_Col.OpenDocTotali "1=0"
    If ClsRegDoc_Col.RegDocIn.GestDocTestaProge Then
        ClsRegDoc_Col.OpenDocTestaProge "1=0"
    End If
    ClsRegDoc_Col.OpenDocCorpo "1=0"
    ClsRegDoc_Col.OpenDocCorpoOrd "1=0"
    If ClsRegDoc_Col.RegDocIn.GestDocCorpoProge Then
        ClsRegDoc_Col.OpenDocCorpoProge "1=0"
    End If
    If Pbol_MovLotti Then
        ClsRegDoc_Col.OpenDocCorpoLot "1=0"
    End If
    ClsRegDoc_Col.OpenControPartite "1=0"
    ClsRegDoc_Col.OpenCastIva "1=0"
    ClsRegDoc_Col.OpenScadenze "1=0"
    '
    ' Invoco il metodo di aggiunta di una testata
    '
    ClsRegDoc_Col.AddNewDocTestata , False, True
    ClsRegDoc_Col.AddNewDocTotali
    ClsRegDoc_Col.AddNewDocDatiAcc
    '
    ' Recupero l'eventuale sezionale fisso agganciato al deposito collegato
    '
    Pstr_Sql = " SELECT" & _
               "    MG58_FLGSEZFIX," & _
               "    (" & _
               "        SELECT TOP 1" & _
               "            MG59_SEZ_CG68" & _
               "        FROM" & _
               "            MG59_DEPNUMFIX WITH (NOLOCK)" & _
               "        WHERE" & _
               "            MG59_DITTA_CG18 = MG58_DITTA_CG18 AND" & _
               "            MG59_CODDEP = MG58_CODDEP" & _
               "    ) AS MG59_SEZ_CG68" & _
               " FROM" & _
               "    MG58_DEPOSITI WITH (NOLOCK)" & _
               " WHERE" & _
               "    MG58_DITTA_CG18 = " & CodiceDitta & " AND" & _
               "    MG58_CODDEP = '" & ClsRegDoc.rstDocTestata("DO11_CODDEPCOL").value & "'"
    Set Prst_SezDefault = Gcon_Connect.Execute(Pstr_Sql, , adCmdText)
    If Prst_SezDefault.EOF Then
        Pstr_SezDefault = ""
    Else
        Pstr_SezDefault = IIf(CDecN(Prst_SezDefault("MG58_FLGSEZFIX").value) = 1, RTrimN(Prst_SezDefault("MG59_SEZ_CG68").value), "")
    End If
    Prst_SezDefault.Close
    Set Prst_SezDefault = Nothing
    '
    ' Verifico che il sezionale sia impostato
    '
    If Pstr_SezDefault = "" Then
        If RTrimN(ClsRegDoc_Col.rstDocTestata("DO11_SEZDOC").value) = "" Then
            ClsRegDoc_Col.CambiaDocTestataSezDoc ClsRegDoc.rstDocTestata("DO11_SEZDOC").value
        End If
    Else
        If RTrimN(ClsRegDoc_Col.rstDocTestata("DO11_SEZDOC").value) <> Pstr_SezDefault Then
            ClsRegDoc_Col.CambiaDocTestataSezDoc Pstr_SezDefault
        End If
    End If
    '
    ' Imposto i dati di testata
    '
    For Each ADOField In ClsRegDoc_Col.rstDocTestata.Fields
        Select Case ADOField.Name
            Case "DO11_DITTA_CG18", "DO11_NUMREG_CO99", "DO11_DOCUM_MG36", "DO11_TIPODOC", "DO11_STIPODOC"
            Case "DO11_SEZDOC", "DO11_FLGDOCBIS", "DO11_TIPODOCNUMERAZ", "DO11_INDTIPOCODCF", "DO11_LASTCHANGE"
            Case "DO11_NUMDOC"
                If RTrimN(ADOField.value) = "" Then
                    ADOField.value = ClsRegDoc.rstDocTestata(ADOField.Name).value
                End If
            Case "DO11_CAUSMAG_MG51"
            Case "DO11_NUMREGCOL_CO99"
                ADOField.value = ClsRegDoc.rstDocTestata("DO11_NUMREG_CO99").value
            Case "DO11_CODDEP"
                ADOField.value = ClsRegDoc.rstDocTestata("DO11_CODDEPCOL").value
            Case "DO11_CODDEPCOL"
                ADOField.value = ClsRegDoc.rstDocTestata("DO11_CODDEP").value
            Case "DO11_INDGENERDOC"
                If CDecN(ClsRegDoc_Col.rstDocTestata("DO11_INDTIPOCODCF").value) = 0 And CDecN(ClsRegDoc.rstDocTestata("DO11_INDTIPOCODCF").value) <> 0 Then
                    ADOField.value = 2
                End If
            Case "DO11_DITTACF_CG44", "DO11_TIPOCF_CG44", "DO11_CLIFOR_CG44"
                If CDecN(ClsRegDoc_Col.rstDocTestata("DO11_INDTIPOCODCF").value) <> 0 Then
                    If Not IsNull(ClsRegDoc.rstDocTestata(ADOField.Name).value) Then
                        ADOField.value = ClsRegDoc.rstDocTestata(ADOField.Name).value
                    End If
                End If
            ' <Marchionni 2011.09.27 - Correzione relativa al campo Guid>
            Case "DO11_GUID"
                ClsRegDoc_Col.BoStatoTestata.DuplicaStato ADOField.value, ClsRegDoc.rstDocTestata(ADOField.Name).value
                'TODO: Aggiungere la duplica degli attributi
            '</Marchionni 2011.09.27 - Correzione relativa al campo Guid>
                
            Case Else
                If RTrimN(ADOField.value) <> RTrimN(ClsRegDoc.rstDocTestata(ADOField.Name).value) Then
                    ADOField.value = ClsRegDoc.rstDocTestata(ADOField.Name).value
                End If
        End Select
    Next
    '
    ' Imposto tutti i campi dei totali
    '
    For Each ADOField In ClsRegDoc_Col.rstDocTotali.Fields
        Select Case ADOField.Name
            Case "DO13_DITTA_CG18", "DO13_NUMREG_CO99"
            Case Else
                If RTrimN(ADOField.value) <> RTrimN(ClsRegDoc.rstDocTotali(ADOField.Name).value) Then
                    ADOField.value = ClsRegDoc.rstDocTotali(ADOField.Name).value
                End If
        End Select
    Next
    '
    ' Imposto tutti i dati di accompagnamento
    '
    For Each ADOField In ClsRegDoc_Col.rstDocDatiAcc.Fields
        Select Case ADOField.Name
            Case "DO14_DITTA_CG18", "DO14_NUMREG_CO99"
            Case Else
                If RTrimN(ADOField.value) <> RTrimN(ClsRegDoc.rstDocDatiAcc(ADOField.Name).value) Then
                    ADOField.value = ClsRegDoc.rstDocDatiAcc(ADOField.Name).value
                End If
        End Select
    Next
    '
    ' Imposto gli eventuali dati di testata dei progetti
    '
    
    If ClsRegDoc_Col.RegDocIn.GestDocTestaProge Then
        If ClsRegDoc.rstDocTestaProge.RecordCount > 0 Then
            ClsRegDoc_Col.AddNewDocTestaProge
            '
            ' Valorizzazione progetto/sottoprogetto/nodo di riferimento
            '
            ClsRegDoc_Col.rstDocTestaProge("DO24_CODPROGETTO").value = ClsRegDoc.rstDocTestaProge("DO24_CODPROGCOL").value
            ClsRegDoc_Col.rstDocTestaProge("DO24_CODSPROGETTO").value = ClsRegDoc.rstDocTestaProge("DO24_CODSPROGCOL").value
            ClsRegDoc_Col.rstDocTestaProge("DO24_NODORIF_PD0C").value = ClsRegDoc.rstDocTestaProge("DO24_NODORIFCOL_PD0C").value
            ClsRegDoc_Col.rstDocTestaProge("DO24_CODDIP_PD06").value = ClsRegDoc.rstDocTestaProge("DO24_CODDIP_PD06").value
            '
            ' Valorizzazione progetto/sottoprogetto/nodo di riferimento collegati
            '
            ClsRegDoc_Col.rstDocTestaProge("DO24_CODPROGCOL").value = ClsRegDoc.rstDocTestaProge("DO24_CODPROGETTO").value
            ClsRegDoc_Col.rstDocTestaProge("DO24_CODSPROGCOL").value = ClsRegDoc.rstDocTestaProge("DO24_CODSPROGETTO").value
            ClsRegDoc_Col.rstDocTestaProge("DO24_NODORIFCOL_PD0C").value = ClsRegDoc.rstDocTestaProge("DO24_NODORIF_PD0C").value
            '
            ' Valorizzazione variabile "solo progetto collegato"
            '
            If (RTrimN(ClsRegDoc.rstDocTestaProge("DO24_CODPROGCOL").value) <> "") And (RTrimN(ClsRegDoc.rstDocTestaProge("DO24_CODPROGETTO").value) = "") Then
                Pbol_SoloProgColl = True
            End If
'''            If RTrimN(ClsRegDoc.rstDocTestaProge("DO24_CODPROGCOL").Value) <> "" Then
'''                ClsRegDoc_Col.rstDocTestaProge("DO24_CODPROGETTO").Value = ClsRegDoc.rstDocTestaProge("DO24_CODPROGCOL").Value
'''                ClsRegDoc_Col.rstDocTestaProge("DO24_CODSPROGETTO").Value = ClsRegDoc.rstDocTestaProge("DO24_CODSPROGCOL").Value
'''                If RTrimN(ClsRegDoc.rstDocTestaProge("DO24_NODORIFCOL_PD0C").Value) <> "" Then
'''                    ClsRegDoc_Col.rstDocTestaProge("DO24_NODORIF_PD0C").Value = ClsRegDoc.rstDocTestaProge("DO24_NODORIFCOL_PD0C").Value
'''                End If
'''                If RTrimN(ClsRegDoc.rstDocTestaProge("DO24_CODPROGETTO").Value) <> "" Then
'''                    ClsRegDoc_Col.rstDocTestaProge("DO24_CODPROGCOL").Value = ClsRegDoc.rstDocTestaProge("DO24_CODPROGETTO").Value
'''                    ClsRegDoc_Col.rstDocTestaProge("DO24_CODSPROGCOL").Value = ClsRegDoc.rstDocTestaProge("DO24_CODSPROGETTO").Value
'''                    If RTrimN(ClsRegDoc.rstDocTestaProge("DO24_NODORIF_PD0C").Value) <> "" Then
'''                        ClsRegDoc_Col.rstDocTestaProge("DO24_NODORIFCOL_PD0C").Value = ClsRegDoc.rstDocTestaProge("DO24_NODORIF_PD0C").Value
'''                    End If
'''                Else
'''                    Pbol_SoloProgColl = True
'''                End If
'''            Else
'''                If RTrimN(ClsRegDoc.rstDocTestaProge("DO24_CODPROGETTO").Value) <> "" Then
'''                    ClsRegDoc_Col.rstDocTestaProge("DO24_CODPROGETTO").Value = ClsRegDoc.rstDocTestaProge("DO24_CODPROGETTO").Value
'''                    ClsRegDoc_Col.rstDocTestaProge("DO24_CODSPROGETTO").Value = ClsRegDoc.rstDocTestaProge("DO24_CODSPROGETTO").Value
'''                    If RTrimN(ClsRegDoc.rstDocTestaProge("DO24_NODORIF_PD0C").Value) <> "" Then
'''                        ClsRegDoc_Col.rstDocTestaProge("DO24_NODORIF_PD0C").Value = ClsRegDoc.rstDocTestaProge("DO24_NODORIF_PD0C").Value
'''                    End If
'''                End If
'''            End If
'</Luigi Att. 86353 ERRORE IN GENERAZIONE DOCUMENTO COLLEGATO>
            If RTrimN(ClsRegDoc.rstDocTestaProge("DO24_CODCOMMESSA").value) <> "" Then
                ClsRegDoc_Col.rstDocTestaProge("DO24_CODCOMMESSA").value = ClsRegDoc.rstDocTestaProge("DO24_CODCOMMESSA").value
                ClsRegDoc_Col.rstDocTestaProge("DO24_CODSCOMMESSA").value = ClsRegDoc.rstDocTestaProge("DO24_CODSCOMMESSA").value
            End If
            ClsRegDoc_Col.UpdateDocTestaProge
        End If
    End If
    '
    ' Ciclo sulle righe di corpo
    '
    Set rst_DocCorpo_Clone = ClsRegDoc.rstDocCorpo.Clone
    Set rst_DocCorpoOrd_Clone = ClsRegDoc.rstDocCorpoOrd.Clone
    If ClsRegDoc_Col.RegDocIn.GestDocCorpoProge Then
        Set rst_DocCorpoProge_Clone = ClsRegDoc.rstDocCorpoProge.Clone
    End If
    If Pbol_MovLotti Then
        Set rst_DocCorpoLot_Clone = ClsRegDoc.rstDocCorpoLot.Clone
    End If
    rst_DocCorpo_Clone.MoveFirst
    While Not rst_DocCorpo_Clone.EOF
        ClsRegDoc_Col.AddNewDocCorpo , , , False
        For Each ADOField In ClsRegDoc_Col.rstDocCorpo.Fields
            Select Case ADOField.Name
                Case "DO30_DITTA_CG18", "DO30_NUMREG_CO99"
                Case "DO30_CAUSMAG_MG51"
                Case "DO30_CODDEP_MG58"
                    ADOField.value = rst_DocCorpo_Clone("DO30_CODDEPCOL_MG58").value
                Case "DO30_CODDEPCOL_MG58"
                    ADOField.value = rst_DocCorpo_Clone("DO30_CODDEP_MG58").value
                Case "DO30_OPZIONE_MG5E"
                    If RTrimN(ADOField.value) <> RTrimN(rst_DocCorpo_Clone(ADOField.Name).value) Then
                        ADOField.value = rst_DocCorpo_Clone(ADOField.Name).value
                    Else
                        If CDecN(rst_DocCorpo_Clone("DO30_INDTIPORIGA").value) = 0 And RTrimN(rst_DocCorpo_Clone(ADOField.Name).value) = "" And IsNull(ADOField.value) Then
                            ADOField.value = ""
                        End If
                    End If
                    ' <Marchionni 2011.09.27 - Correzione relativa al campo Guid>
                Case "DO30_GUID"
                    ClsRegDoc_Col.BoStatoCorpo.DuplicaStato ADOField.value, rst_DocCorpo_Clone(ADOField.Name).value
                    ' </Marchionni 2011.09.27 - Correzione relativa al campo Guid>
                
                Case Else
                    If RTrimN(ADOField.value) <> RTrimN(rst_DocCorpo_Clone(ADOField.Name).value) Then
                        ADOField.value = rst_DocCorpo_Clone(ADOField.Name).value
                    End If
            End Select
        Next
        ClsRegDoc_Col.AddNewDocCorpoOrd
        rst_DocCorpoOrd_Clone.Filter = "DO31_PROGRIGA = " & CDecN(rst_DocCorpo_Clone("DO30_PROGRIGA").value)
        If Not rst_DocCorpoOrd_Clone.EOF Then
            For Each ADOField In ClsRegDoc_Col.rstDocCorpoOrd.Fields
                Select Case ADOField.Name
                    Case "DO31_DITTA_CG18", "DO31_NUMREG_CO99", "DO31_PROGRIGA"
                    Case "DO31_QTA1CONS", "DO31_QTA2CONS", "DO31_COLLICONS", "DO31_INDSTATODESCSP", "DO31_FLGNONPIUEV", "DO31_INDSTATOCONS", "DO31_QTA1TRASFP", "DO31_QTA2TRASFP", "DO31_COLLITRASFP", "DO31_QTA1TRASFDDTCL", "DO31_QTA2TRASFDDTCL", "DO31_COLLITRASFDDTCL"
                    Case Else
                        If RTrimN(ADOField.value) <> RTrimN(rst_DocCorpoOrd_Clone(ADOField.Name).value) Then
                            ADOField.value = rst_DocCorpoOrd_Clone(ADOField.Name).value
                        End If
                End Select
            Next
            ClsRegDoc_Col.UpdateDocCorpoOrd
        End If
        If ClsRegDoc_Col.RegDocIn.GestDocCorpoProge Then
            rst_DocCorpoProge_Clone.Filter = "DO41_PROGRIGA = " & CDecN(rst_DocCorpo_Clone("DO30_PROGRIGA").value)
            If rst_DocCorpoProge_Clone.RecordCount > 0 Then
                ClsRegDoc_Col.AddNewDocCorpoProge
                '
                ' Valorizzazione progetto / sottoprogetto / nodo di riferimento
                '
                ClsRegDoc_Col.rstDocCorpoProge("DO41_CODPROGETTO").value = rst_DocCorpoProge_Clone("DO41_CODPROGCOL").value
                ClsRegDoc_Col.rstDocCorpoProge("DO41_CODSPROGETTO").value = rst_DocCorpoProge_Clone("DO41_CODSPROGCOL").value
                ClsRegDoc_Col.rstDocCorpoProge("DO41_NODORIF_PD0C").value = rst_DocCorpoProge_Clone("DO41_NODORIFCOL_PD0C").value
                '
                ' Valorizzazione progetto / sottoprogetto / nodo di riferimento collegato
                '
                ClsRegDoc_Col.rstDocCorpoProge("DO41_CODPROGCOL").value = rst_DocCorpoProge_Clone("DO41_CODPROGETTO").value
                ClsRegDoc_Col.rstDocCorpoProge("DO41_CODSPROGCOL").value = rst_DocCorpoProge_Clone("DO41_CODSPROGETTO").value
                ClsRegDoc_Col.rstDocCorpoProge("DO41_NODORIFCOL_PD0C").value = rst_DocCorpoProge_Clone("DO41_NODORIF_PD0C").value
                '
                ' Valorizzazione variabile solo progetto collegato
                '
                If (RTrimN(rst_DocCorpoProge_Clone("DO41_CODPROGCOL").value) <> "") And (RTrimN(rst_DocCorpoProge_Clone("DO41_CODPROGETTO").value) = "") Then
                    Pbol_SoloProgColl = True
                End If
'''                If RTrimN(rst_DocCorpoProge_Clone("DO41_CODPROGCOL").Value) <> "" Then
'''                    ClsRegDoc_Col.rstDocCorpoProge("DO41_CODPROGETTO").Value = rst_DocCorpoProge_Clone("DO41_CODPROGCOL").Value
'''                    ClsRegDoc_Col.rstDocCorpoProge("DO41_CODSPROGETTO").Value = rst_DocCorpoProge_Clone("DO41_CODSPROGCOL").Value
'''                    If RTrimN(rst_DocCorpoProge_Clone("DO41_NODORIFCOL_PD0C").Value) <> "" Then
'''                        '<Luigi 23/12/2009 att. 86073>
'''                        ClsRegDoc_Col.rstDocCorpoProge("DO41_NODORIF_PD0C").Value = rst_DocCorpoProge_Clone("DO41_NODORIFCOL_PD0C").Value
'''                        '</Luigi 23/12/2009 att. 86073>
'''                    End If
'''                    If RTrimN(rst_DocCorpoProge_Clone("DO41_CODPROGETTO").Value) <> "" Then
'''                        ClsRegDoc_Col.rstDocCorpoProge("DO41_CODPROGCOL").Value = rst_DocCorpoProge_Clone("DO41_CODPROGETTO").Value
'''                        ClsRegDoc_Col.rstDocCorpoProge("DO41_CODSPROGCOL").Value = rst_DocCorpoProge_Clone("DO41_CODSPROGETTO").Value
'''                        If RTrimN(rst_DocCorpoProge_Clone("DO41_NODORIF_PD0C").Value) <> "" Then
'''                            ClsRegDoc_Col.rstDocCorpoProge("DO41_NODORIFCOL_PD0C").Value = rst_DocCorpoProge_Clone("DO41_NODORIF_PD0C").Value
'''                        End If
'''                    Else
'''                        Pbol_SoloProgColl = True
'''                    End If
'''                Else
'''                    If RTrimN(rst_DocCorpoProge_Clone("DO41_CODPROGETTO").Value) <> "" Then
'''                        ClsRegDoc_Col.rstDocCorpoProge("DO41_CODPROGETTO").Value = rst_DocCorpoProge_Clone("DO41_CODPROGETTO").Value
'''                        ClsRegDoc_Col.rstDocCorpoProge("DO41_CODSPROGETTO").Value = rst_DocCorpoProge_Clone("DO41_CODSPROGETTO").Value
'''                        If RTrimN(rst_DocCorpoProge_Clone("DO41_NODORIF_PD0C").Value) <> "" Then
'''                            ClsRegDoc_Col.rstDocCorpoProge("DO41_NODORIF_PD0C").Value = rst_DocCorpoProge_Clone("DO41_NODORIF_PD0C").Value
'''                        End If
'''                    End If
'''                End If
'</Luigi Att. 86353 ERRORE IN GENERAZIONE DOCUMENTO COLLEGATO>
                If RTrimN(rst_DocCorpoProge_Clone("DO41_CODCOMMESSA").value) <> "" Then
                    ClsRegDoc_Col.rstDocCorpoProge("DO41_CODCOMMESSA").value = rst_DocCorpoProge_Clone("DO41_CODCOMMESSA").value
                    ClsRegDoc_Col.rstDocCorpoProge("DO41_CODSCOMMESSA").value = rst_DocCorpoProge_Clone("DO41_CODSCOMMESSA").value
                End If
                If RTrimN(rst_DocCorpoProge_Clone("DO41_NODO_PD0C").value) <> "" Then
                    ClsRegDoc_Col.rstDocCorpoProge("DO41_NODO_PD0C").value = rst_DocCorpoProge_Clone("DO41_NODO_PD0C").value
                End If
                If RTrimN(rst_DocCorpoProge_Clone("DO41_ATTIVITA_PD0D").value) <> "" Then
                    ClsRegDoc_Col.rstDocCorpoProge("DO41_ATTIVITA_PD0D").value = rst_DocCorpoProge_Clone("DO41_ATTIVITA_PD0D").value
                End If
                If RTrimN(rst_DocCorpoProge_Clone("DO41_SPESA_PD64").value) <> "" Then
                    ClsRegDoc_Col.rstDocCorpoProge("DO41_SPESA_PD64").value = rst_DocCorpoProge_Clone("DO41_SPESA_PD64").value
                End If
                If RTrimN(rst_DocCorpoProge_Clone("DO41_CODDIP_PD06").value) <> "" Then
                    ClsRegDoc_Col.rstDocCorpoProge("DO41_CODDIP_PD06").value = rst_DocCorpoProge_Clone("DO41_CODDIP_PD06").value
                End If
                ClsRegDoc_Col.UpdateDocCorpoProge
            End If
        End If
        If Pbol_MovLotti Then
            Pstr_Filter = "DO52_PROGRIGA = " & CDecN(rst_DocCorpo_Clone("DO30_PROGRIGA").value)
            If Pint_Num_Filter_MG4F > 1 Then
                Pstr_Filter_MG4F = ""
                For Pint_Index = 1 To Pint_Num_Filter_MG4F
                    Pstr_Filter_MG4F = Pstr_Filter_MG4F & IIf(Pint_Index > 1, " OR ", "") & "(" & Pstr_Filter & " AND DO52_PROG_MG4F = " & Parr_Filter_MG4F(Pint_Index) & ")"
                Next Pint_Index
            End If
            If Pbol_DO52_Div Then
                If Pint_Num_Filter_MG4F > 1 Then
                    rst_DocCorpoLot_Clone.Filter = Pstr_Filter_MG4F
                Else
                    rst_DocCorpoLot_Clone.Filter = Pstr_Filter & " AND " & Pstr_Filter_MG4F
                End If
            Else
                rst_DocCorpoLot_Clone.Filter = Pstr_Filter
                Plng_NumRec52_Orig = rst_DocCorpoLot_Clone.RecordCount
                If Pint_Num_Filter_MG4F > 1 Then
                    rst_DocCorpoLot_Clone.Filter = Pstr_Filter_MG4F
                Else
                    rst_DocCorpoLot_Clone.Filter = Pstr_Filter & " AND " & Pstr_Filter_MG4F
                End If
                Plng_NumRec52_Filter = rst_DocCorpoLot_Clone.RecordCount
                If Plng_NumRec52_Filter <> Plng_NumRec52_Orig Then
                    Pbol_DO52_Div = True
                End If
            End If
            While Not rst_DocCorpoLot_Clone.EOF
                ClsRegDoc_Col.AddNewDocCorpoLot
                For Each ADOField In ClsRegDoc_Col.rstDocCorpoLot.Fields
                    Select Case ADOField.Name
                        Case "DO52_DITTA_CG18", "DO52_NUMREG_CO99", "DO52_PROGRIGA", "DO52_NUMREGRIF_CO99", "DO52_PROGRIGARIF_DO30", "DO52_PROGRIF_DO52"
                        Case "DO52_CAUSMAG_MG51"
'                            If GIntCausMagDocCollegato > 0 Then
'                                ADOField.value = GIntCausMagDocCollegato
'                            End If
                        Case "DO52_CODDEP_MG58"
                            ADOField.value = rst_DocCorpo_Clone("DO30_CODDEPCOL_MG58").value
                        Case "DO52_CODPROG_PD14"
                            If Not IsNull(rst_DocCorpoLot_Clone(ADOField.Name).value) Then
                                Pvar_Progetto = rst_DocCorpoLot_Clone(ADOField.Name).value
                                If ClsRegDoc_Col.RegDocIn.GestDocCorpoProge Then
                                    If Not ClsRegDoc_Col.rstDocCorpoProge.EOF And Not ClsRegDoc_Col.rstDocCorpoProge.BOF Then
                                        If ClsRegDoc_Col.rstDocCorpoProge("DO41_PROGRIGA").value = ClsRegDoc_Col.rstDocCorpo("DO30_PROGRIGA").value Then
                                            If Not IsNull(ClsRegDoc_Col.rstDocCorpoProge("DO41_CODPROGETTO").value) Then
                                                Pvar_Progetto = ClsRegDoc_Col.rstDocCorpoProge("DO41_CODPROGETTO").value
                                            End If
                                        End If
                                    End If
                                End If
                                ADOField.value = Pvar_Progetto
                            End If
                        Case Else
                            If Not IsNull(rst_DocCorpoLot_Clone(ADOField.Name).value) Then
                                ADOField.value = rst_DocCorpoLot_Clone(ADOField.Name).value
                            End If
                    End Select
                Next
                ClsRegDoc_Col.UpdateDocCorpoLot
                rst_DocCorpoLot_Clone.MoveNext
            Wend
        End If
        ClsRegDoc_Col.UpdateDocCorpo
        rst_DocCorpo_Clone.MoveNext
    Wend
    '
    ' Se diversità di parametrizzazione lotti o in presenza del solo progetto collegato
    ' metto il flag di documento non modificabile
    '
    If Pbol_DO52_Div Or Pbol_SoloProgColl Then
        ClsRegDoc_Col.rstDocTestata("DO11_INDGENERDOC").value = 2
    End If
    '
    ' Registrazione del documento collegato
    '
    ClsRegDoc_Col.UpdateDocTestata
    '
    ' Restituisco il numero di registrazione
    '
    GeneraDocumentoCollegato = ClsRegDoc_Col.rstDocTestata("DO11_NUMREG_CO99").value
    
   
    
    '
    ' Rilascio i recordsets
    '
    Set rst_DocCorpo_Clone = Nothing
    Set rst_DocCorpoOrd_Clone = Nothing
    Set rst_DocCorpoProge_Clone = Nothing
    Set rst_DocCorpoLot_Clone = Nothing
    '
    ' Rilascio il Bo del documento collegato
    '
    ClsRegDoc_Col.Terminate
    Set ClsRegDoc_Col.RegDocIn.ActiveInterface = Nothing
    Set ClsRegDoc_Col.RegDocIn.Connection = Nothing
    Set ClsRegDoc_Col.RegDocIn = Nothing
    Set ClsRegDoc_Col = Nothing
End Function


