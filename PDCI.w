&ANALYZE-SUSPEND _VERSION-NUMBER AB_v10r12 GUI
&ANALYZE-RESUME
&Scoped-define WINDOW-NAME C-Win
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _DEFINITIONS C-Win 
/*------------------------------------------------------------------------

  File: 

  Description: 

  Input Parameters:
      <none>

  Output Parameters:
      <none>

  Author: 

  Created: 

------------------------------------------------------------------------*/
/*          This .W file was created with the Progress AppBuilder.      */
/*----------------------------------------------------------------------*/

/* Create an unnamed pool to store all the widgets created 
     by this procedure. This is a good default which assures
     that this procedure's triggers and internal procedures 
     will execute in this procedure's storage, and that proper
     cleanup will occur on deletion of the procedure. */

CREATE WIDGET-POOL.

/* ***************************  Definitions  ************************** */

/* Parameters Definitions ---                                           */

/* Local Variable Definitions ---                                       */

DEF VAR cFile AS CHAR NO-UNDO.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&ANALYZE-SUSPEND _UIB-PREPROCESSOR-BLOCK 

/* ********************  Preprocessor Definitions  ******************** */

&Scoped-define PROCEDURE-TYPE Window
&Scoped-define DB-AWARE no

/* Name of designated FRAME-NAME and/or first browse and/or first query */
&Scoped-define FRAME-NAME DEFAULT-FRAME

/* Standard List Definitions                                            */
&Scoped-Define ENABLED-OBJECTS RECT-1 RECT-2 edt-endereco BUTTON-1 ~
edt-pedido chk-visualizar btn-Imprimir 
&Scoped-Define DISPLAYED-OBJECTS edt-endereco edt-pedido chk-visualizar 

/* Custom List Definitions                                              */
/* List-1,List-2,List-3,List-4,List-5,List-6                            */

/* _UIB-PREPROCESSOR-BLOCK-END */
&ANALYZE-RESUME



/* ***********************  Control Definitions  ********************** */

/* Define the widget handle for the window                              */
DEFINE VAR C-Win AS WIDGET-HANDLE NO-UNDO.

/* Definitions of the field level widgets                               */
DEFINE BUTTON btn-Imprimir 
     LABEL "Executar" 
     SIZE 15 BY 1.13
     FONT 4.

DEFINE BUTTON BUTTON-1 
     IMAGE-UP FILE "adeicon/open.bmp":U
     LABEL "" 
     SIZE 6 BY 1.13.

DEFINE VARIABLE edt-endereco AS CHARACTER FORMAT "X(256)":U 
     LABEL "Salvar Em" 
     VIEW-AS FILL-IN 
     SIZE 46 BY 1 NO-UNDO.

DEFINE VARIABLE edt-pedido AS INTEGER FORMAT "->,>>>,>>9":U INITIAL 0 
     LABEL "Pedido" 
     VIEW-AS FILL-IN 
     SIZE 22 BY 1 NO-UNDO.

DEFINE RECTANGLE RECT-1
     EDGE-PIXELS 2 GRAPHIC-EDGE  NO-FILL   
     SIZE 67 BY 2.5.

DEFINE RECTANGLE RECT-2
     EDGE-PIXELS 2 GRAPHIC-EDGE  NO-FILL   
     SIZE 67 BY 2.25.

DEFINE VARIABLE chk-visualizar AS LOGICAL INITIAL no 
     LABEL "Visualizar Arquivo" 
     VIEW-AS TOGGLE-BOX
     SIZE 15.43 BY .83
     FONT 4 NO-UNDO.


/* ************************  Frame Definitions  *********************** */

DEFINE FRAME DEFAULT-FRAME
     edt-endereco AT ROW 2 COL 12 COLON-ALIGNED WIDGET-ID 12
     BUTTON-1 AT ROW 2 COL 61 WIDGET-ID 14
     edt-pedido AT ROW 4.5 COL 9 COLON-ALIGNED WIDGET-ID 6
     chk-visualizar AT ROW 4.75 COL 53 WIDGET-ID 16
     btn-Imprimir AT ROW 6.5 COL 54 WIDGET-ID 2
     RECT-1 AT ROW 1.25 COL 2 WIDGET-ID 8
     RECT-2 AT ROW 4 COL 2 WIDGET-ID 10
    WITH 1 DOWN NO-BOX KEEP-TAB-ORDER OVERLAY 
         SIDE-LABELS NO-UNDERLINE THREE-D 
         AT COL 1 ROW 1
         SIZE 68.86 BY 6.88 WIDGET-ID 100.


/* *********************** Procedure Settings ************************ */

&ANALYZE-SUSPEND _PROCEDURE-SETTINGS
/* Settings for THIS-PROCEDURE
   Type: Window
   Allow: Basic,Browse,DB-Fields,Window,Query
   Other Settings: COMPILE
 */
&ANALYZE-RESUME _END-PROCEDURE-SETTINGS

/* *************************  Create Window  ************************** */

&ANALYZE-SUSPEND _CREATE-WINDOW
IF SESSION:DISPLAY-TYPE = "GUI":U THEN
  CREATE WINDOW C-Win ASSIGN
         HIDDEN             = YES
         TITLE              = "PDCI - Imprimir Pedido"
         HEIGHT             = 6.88
         WIDTH              = 68.86
         MAX-HEIGHT         = 16
         MAX-WIDTH          = 80
         VIRTUAL-HEIGHT     = 16
         VIRTUAL-WIDTH      = 80
         RESIZE             = yes
         SCROLL-BARS        = no
         STATUS-AREA        = no
         BGCOLOR            = ?
         FGCOLOR            = ?
         KEEP-FRAME-Z-ORDER = yes
         THREE-D            = yes
         MESSAGE-AREA       = no
         SENSITIVE          = yes.
ELSE {&WINDOW-NAME} = CURRENT-WINDOW.
/* END WINDOW DEFINITION                                                */
&ANALYZE-RESUME



/* ***********  Runtime Attributes and AppBuilder Settings  *********** */

&ANALYZE-SUSPEND _RUN-TIME-ATTRIBUTES
/* SETTINGS FOR WINDOW C-Win
  VISIBLE,,RUN-PERSISTENT                                               */
/* SETTINGS FOR FRAME DEFAULT-FRAME
   FRAME-NAME                                                           */
IF SESSION:DISPLAY-TYPE = "GUI":U AND VALID-HANDLE(C-Win)
THEN C-Win:HIDDEN = no.

/* _RUN-TIME-ATTRIBUTES-END */
&ANALYZE-RESUME

 



/* ************************  Control Triggers  ************************ */

&Scoped-define SELF-NAME C-Win
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL C-Win C-Win
ON END-ERROR OF C-Win /* PDCI - Imprimir Pedido */
OR ENDKEY OF {&WINDOW-NAME} ANYWHERE DO:
  /* This case occurs when the user presses the "Esc" key.
     In a persistently run window, just ignore this.  If we did not, the
     application would exit. */
  IF THIS-PROCEDURE:PERSISTENT THEN RETURN NO-APPLY.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL C-Win C-Win
ON WINDOW-CLOSE OF C-Win /* PDCI - Imprimir Pedido */
DO:
  /* This event will close the window and terminate the procedure.  */
  APPLY "CLOSE":U TO THIS-PROCEDURE.
  RETURN NO-APPLY.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME btn-Imprimir
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL btn-Imprimir C-Win
ON CHOOSE OF btn-Imprimir IN FRAME DEFAULT-FRAME /* Executar */
DO:
  
DEF VAR ex AS COM-HANDLE NO-UNDO.
DEF VAR ChWorkSheet AS COM-HANDLE NO-UNDO.


CREATE "excel.application" ex.

ex:VISIBLE = FALSE.
ex:workbooks:ADD("\\server03\ERP\_custom\_brandl\pedido.xlt").

/* ----- AREA DE CONTEUDO ----- */

DEF VAR pedido AS INT.

ASSIGN pedido = int(edt-pedido:SCREEN-VALUE IN FRAME DEFAULT-FRAME).

FOR FIRST pedido-compr WHERE num-pedido = pedido,

    //Rela»’o Transporte
    FIRST transporte NO-LOCK WHERE transporte.cod-transp = pedido-compr.cod-transp, 
    //Rela»’o Condi»’o de Pagamento
    FIRST cond-pagto NO-LOCK WHERE cond-pagto.cod-cond-pag = pedido-compr.cod-cond-pag,                
    //Rela»’o Codigo do Emitente
    FIRST emitente NO-LOCK WHERE emitente.cod-emitente = pedido-compr.cod-emitente,
    //Rela»’o Estabelecimento
    FIRST estabelec NO-LOCK WHERE estabelec.cod-estabel = pedido-compr.end-entrega.
    
    //BM
    FOR EACH ex-ordem-compra NO-LOCK WHERE ex-ordem-compra.num-pedido = pedido-compr.num-pedido.
    
        //BM do Pedido       
        ex:range("N8"):VALUE = ex-ordem-compra.numero-bm.

    END.
    

    //Numero do Pedido de Compra
    ex:range("F8"):VALUE = pedido-compr.num-pedido. 

        
    //Data do Pedido de Compra
    ex:range("AE8"):VALUE = pedido-compr.data-pedido.


    //Natureza de Operacao
    IF pedido-compr.natureza = 0 THEN
        ex:range("A11"):VALUE = "Servi»o".
    ELSE IF pedido-compr.natureza = 1  THEN
        ex:range("A11"):VALUE = "Compra".
    ELSE IF pedido-compr.natureza = 2 THEN
        ex:range("A11"):VALUE = "Beneficiamento".

    //Condi»’o de pagamento
    ex:range("H11"):VALUE = STRING(cond-pagto.descricao).
    
    //Transporte
    ex:range("P11"):VALUE = STRING(pedido-compr.cod-transp) + " - " + STRING(transporte.nome). 

    //Tipo de Frete
    IF pedido-compr.frete = 1 THEN
        ex:range("AF11"):VALUE = "Pago".
    ELSE IF pedido-compr.frete = 2  THEN
        ex:range("AF11"):VALUE = "A Pagar".

    /*------ DADOS DE FORNECEDOR ----------- */

    //Codigo do Fornecedor
    ex:range("E13"):VALUE = STRING(pedido-compr.cod-emitente).
    //Razao Social do Fornecedor
    ex:range("A14"):VALUE = STRING(emitente.nome-emit).
    //CNPF Fornecedor
    ex:range("C15"):VALUE = STRING(emitente.cgc).
    //IE Fornecedor
    ex:range("K15"):VALUE = STRING(emitente.ins-estadual).
    //IM Fornecedor
    ex:range("S15"):VALUE = STRING(emitente.ins-municipal).


    //Endere»o Fornecedor
    ex:range("D16"):VALUE = STRING(emitente.endereco).
    //bairro Fronecedor
    ex:range("D17"):VALUE = STRING(emitente.bairro).
    //Cidade Fornecedor
    ex:range("D18"):VALUE = STRING(emitente.cidade).
    //CEP Fornecedor
    ex:range("D19"):VALUE = STRING(emitente.cep).
    //Pais Fornecedor
    //ex:range("AD19"):VALUE = STRING(emitente.pais).
    //Estado Fornecedor
    ex:range("D20"):VALUE = STRING(emitente.estado).

    //Caixa Postal
    ex:range("AB17"):VALUE = STRING(emitente.caixa-postal).
    //Telefone
    ex:range("P17"):VALUE = emitente.telefone.
    //Ramal
    ex:range("P18"):VALUE = emitente.ramal.
    //FAX
    ex:range("P19"):VALUE = STRING(emitente.telefax).
    //TELEX
    ex:range("P20"):VALUE = STRING(emitente.telex).

    /*----------- FIM DADOS DE FORNECEDOR-------------------------*/

    ex:range("A31"):VALUE = STRING(pedido-compr.comentarios).


    /*-------FIM DADOS DE ENTREGA-------------*/


    /* *********** ITENS ************** */
    DEF VAR contadorInsert AS INT.
    DEF VAR contadorItem AS INT.
    ASSIGN contadorInsert = 26.
    ASSIGN contadorItem = 25.

    FOR EACH ordem-compra WHERE ordem-compra.num-pedido = pedido-compr.num-pedido,

        //Relacao Descricao dos itens
        FIRST ITEM WHERE ITEM.it-codigo = ordem-compra.it-codigo.
        
        
        //Codigo do Item
        ex:range("A" + STRING(contadorItem)):VALUE = ordem-compra.it-codigo.
        
        //Descricao do item
        ex:range("D" + STRING(contadorItem)):VALUE = string(item.desc-item + " - " + ordem-compra.narrativa).
        
        IF ordem-compra.codigo-icm = 1 THEN
            ex:range("S" + STRING(contadorItem)):VALUE = "CONS.".
        ELSE IF ordem-compra.codigo-icm = 2 THEN
            ex:range("S" + STRING(contadorItem)):VALUE = "IND.".

        ex:range("W" + STRING(contadorItem)):VALUE = string(item.un).
        ex:range("U" + STRING(contadorItem)):VALUE = ordem-compra.qt-solic.
        ex:range("Y" + STRING(contadorItem)):VALUE = ordem-compra.preco-unit.
        ex:range("AB" + STRING(contadorItem)):VALUE = ordem-compra.valor-descto.

        IF ordem-compra.situacao = 1 THEN
            ex:range("AH" + STRING(contadorItem)):VALUE = "Nao Confirmado".
        ELSE IF ordem-compra.situacao = 2 THEN
            ex:range("AH" + STRING(contadorItem)):VALUE = "Confirmado".
        ELSE IF ordem-compra.situacao = 3 THEN
            ex:range("AH" + STRING(contadorItem)):VALUE = "Cotada".
        ELSE IF ordem-compra.situacao = 4 THEN
            ex:range("AH" + STRING(contadorItem)):VALUE = "Eliminado".
        ELSE IF ordem-compra.situacao = 5 THEN
            ex:range("AH" + STRING(contadorItem)):VALUE = "Em Cotacao".
        ELSE IF ordem-compra.situacao = 6 THEN
            ex:range("AH" + STRING(contadorItem)):VALUE = "Recebido".

        //MESSAGE ordem-compra.preco-unit 

        ex:range("AE" + STRING(contadorItem)):VALUE = DEC(ordem-compra.preco-unit) * (ordem-compra.qt-solic).



        ex:range(string(contadorInsert) + ":" + string(contadorInsert)):COPY().
        ex:range(string(contadorInsert) + ":" + string(contadorInsert)):INSERT().
        ASSIGN contadorInsert = contadorInsert + 1.
        ASSIGN contadorItem = contadorItem + 1.
    END.

    
    DEF VAR arquivo AS CHAR FORMAT "x(50)". 
    ASSIGN arquivo = "Pedido-" + STRING(pedido-compr.num-pedido).
    
    DEF VAR extensao AS CHAR FORMAT "x(50)".
    ASSIGN extensao = ".pdf".

    ex:workbooks:Item(1):ExportAsFixedFormat(0,cFile + "\" + arquivo + extensao,,,,,,,).

    MESSAGE "Arquivo salvo em " + cFile + "\" + arquivo + extensao VIEW-AS ALERT-BOX.

    IF pedido-compr.situacao = 2 THEN
        ASSIGN pedido-compr.situacao = 1.


    IF chk-visualizar:CHECKED = TRUE THEN
        DOS SILENT START VALUE(cFile + "\" + arquivo + extensao).
    
END.

ex:range("A1"):SELECT.

/* ------------------------------------------------------------------------ */

ex:APPLICATION:DisplayAlerts = FALSE.



ex:QUIT().

RELEASE OBJECT ex NO-ERROR.
RELEASE OBJECT chWorksheet NO-ERROR.


END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME BUTTON-1
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL BUTTON-1 C-Win
ON CHOOSE OF BUTTON-1 IN FRAME DEFAULT-FRAME
DO:
  SYSTEM-DIALOG GET-DIR cFile
    TITLE "Escolher uma pasta"
    INITIAL-DIR "C:\Temp".

    edt-endereco:SCREEN-VALUE IN FRAME DEFAULT-FRAME = cFile.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&UNDEFINE SELF-NAME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _MAIN-BLOCK C-Win 


/* ***************************  Main Block  *************************** */

/* Set CURRENT-WINDOW: this will parent dialog-boxes and frames.        */
ASSIGN CURRENT-WINDOW                = {&WINDOW-NAME} 
       THIS-PROCEDURE:CURRENT-WINDOW = {&WINDOW-NAME}.

/* The CLOSE event can be used from inside or outside the procedure to  */
/* terminate it.                                                        */
ON CLOSE OF THIS-PROCEDURE 
   RUN disable_UI.

/* Best default for GUI applications is...                              */
PAUSE 0 BEFORE-HIDE.

/* Now enable the interface and wait for the exit condition.            */
/* (NOTE: handle ERROR and END-KEY so cleanup code will always fire.    */
MAIN-BLOCK:
DO ON ERROR   UNDO MAIN-BLOCK, LEAVE MAIN-BLOCK
   ON END-KEY UNDO MAIN-BLOCK, LEAVE MAIN-BLOCK:
  RUN enable_UI.
  IF NOT THIS-PROCEDURE:PERSISTENT THEN
    WAIT-FOR CLOSE OF THIS-PROCEDURE.
END.

ASSIGN cFile = "C:\Temp".
edt-endereco:SCREEN-VALUE IN FRAME DEFAULT-FRAME = cFile.
chk-visualizar:CHECKED = TRUE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


/* **********************  Internal Procedures  *********************** */

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE disable_UI C-Win  _DEFAULT-DISABLE
PROCEDURE disable_UI :
/*------------------------------------------------------------------------------
  Purpose:     DISABLE the User Interface
  Parameters:  <none>
  Notes:       Here we clean-up the user-interface by deleting
               dynamic widgets we have created and/or hide 
               frames.  This procedure is usually called when
               we are ready to "clean-up" after running.
------------------------------------------------------------------------------*/
  /* Delete the WINDOW we created */
  IF SESSION:DISPLAY-TYPE = "GUI":U AND VALID-HANDLE(C-Win)
  THEN DELETE WIDGET C-Win.
  IF THIS-PROCEDURE:PERSISTENT THEN DELETE PROCEDURE THIS-PROCEDURE.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE enable_UI C-Win  _DEFAULT-ENABLE
PROCEDURE enable_UI :
/*------------------------------------------------------------------------------
  Purpose:     ENABLE the User Interface
  Parameters:  <none>
  Notes:       Here we display/view/enable the widgets in the
               user-interface.  In addition, OPEN all queries
               associated with each FRAME and BROWSE.
               These statements here are based on the "Other 
               Settings" section of the widget Property Sheets.
------------------------------------------------------------------------------*/
  DISPLAY edt-endereco edt-pedido chk-visualizar 
      WITH FRAME DEFAULT-FRAME IN WINDOW C-Win.
  ENABLE RECT-1 RECT-2 edt-endereco BUTTON-1 edt-pedido chk-visualizar 
         btn-Imprimir 
      WITH FRAME DEFAULT-FRAME IN WINDOW C-Win.
  {&OPEN-BROWSERS-IN-QUERY-DEFAULT-FRAME}
  VIEW C-Win.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

