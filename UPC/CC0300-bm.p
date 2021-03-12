
/* Definição dos Parametros Padrão */

DEFINE INPUT PARAMETER p-ind-event  AS CHARACTER NO-UNDO.     /* Evendo do programa Datasul */ 
DEFINE INPUT PARAMETER p-ind-object AS CHARACTER NO-UNDO.     /* Objeto */ 
DEFINE INPUT PARAMETER p-wgh-object AS HANDLE NO-UNDO.        /* Widget do Objeto */ 
DEFINE INPUT PARAMETER p-wgh-frame  AS WIDGET-HANDLE NO-UNDO. /* Frame do Objeto */ 
DEFINE INPUT PARAMETER p-cod-table  AS CHARACTER NO-UNDO.     /* Nome da Tabela */ 
DEFINE INPUT PARAMETER p-row-table  AS ROWID NO-UNDO.         /* Rowid da Tabela */ 

/* Definição de variaveis de fill-in botoes etc... */

DEFINE NEW GLOBAL SHARED VAR wh-tx-lingua AS WIDGET-HANDLE  NO-UNDO.
DEFINE NEW GLOBAL SHARED VAR wh-f-lingua    AS WIDGET-HANDLE  NO-UNDO.


IF p-ind-event = "BEFORE-DISPLAY" AND p-ind-object = "CONTAINER" THEN DO:
    
    /* Cria Texto para usar de Label no campo */
    CREATE TEXT wh-tx-lingua
    ASSIGN FRAME        = p-wgh-frame
          FORMAT       = "x(7)"
          WIDTH        = 7
          SCREEN-VALUE = "BM:"
          ROW          = 5
          COL          = 40.5
          VISIBLE      = YES.

    /* Cria campo do tipo Fill-in */
    CREATE FILL-IN wh-f-lingua
    ASSIGN FRAME   = p-wgh-frame
           NAME               = "Campo"
           DATA-TYPE          = "character"
           FORMAT             = "X(10)"
           WIDTH              = 16
           HEIGHT             = 0.88
           ROW                = 4.85
           COL                = 43.5
           VISIBLE            = YES
           SENSITIVE          = NO.  


       /*procura na tabela es-pais o valor do atributo nome-pais conforme ref
           da ID-ROW da tabela pais*/
       
       FIND FIRST pedido-compr NO-LOCK WHERE ROWID(pedido-compr) = p-row-table NO-ERROR.
            IF AVAIL pedido-compr THEN DO:   /* nome-pais desc-lingua */
                FIND FIRST ex-ordem-compra WHERE ex-ordem-compra.num-pedido = pedido-compr.num-pedido NO-ERROR.

            
           IF AVAIL ex-ordem-compra THEN DO:    
               ASSIGN wh-f-lingua:SCREEN-VALUE = string(ex-ordem-compra.numero-bm).
           END.

           IF NOT AVAIL ex-ordem-compra THEN DO:
               ASSIGN wh-f-lingua:SCREEN-VALUE = "". 
           END.
            
       END.


       END.
