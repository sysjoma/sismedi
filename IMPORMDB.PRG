LOCAL cFile, cCadConex, nConex, nValRet, i, lon, cad, lin, fec, cad2, ;
      num, sex, pes, fla, dini, nOpc

cFile = GETFILE( "mdb" )                  && historia.mdb (access)

IF ( EMPTY( cFile ) )
   RETURN
ENDIF

cCadConex = "DRIVER={Microsoft Access Driver (*.mdb)};" + ;
            "DBQ="+cFile+";"
dini = CTOD( "01/01/1930" )

WAIT WINDOW "Leyendo Archivo, Espere..." NOWAIT

SQLSETPROP( 0, "DispLogin" , 3 )

nConex = SQLSTRINGCONNECT( cCadConex )

IF ( nConex <= 0 )
   AERROR( laErr )
   MESSAGEBOX( "No se pudo conectar con Access. Error: 1"+CHR( 13 )+laErr[ 2 ] )
   RETURN
ENDIF

CREATE TABLE PACIE (HIST_NO C(12),HIST_FECHA D,PACI_TIPO C(30),PACI_APE C(40),PACI_NOM C(40),;
   PACI_LUNAC C(40),PACI_FENAC D,PACI_SEXO C(9),PACI_PESO N(8,3),PACI_FORMU C(30),REFERIDO C(40),;
   PACI_MADRE C(40),PACI_PADRE C(40),PACI_DIREC C(90),PACI_TELFS C(40),HIST_OANTE C(60),HIST_OVACU C(60))
INDEX ON HIST_NO TAG HIST_NO
CREATE TABLE HISTO (HIST_NO C(12),DIEV_FECHA D,DIEV_DES M)
INDEX ON HIST_NO+DTOS( DIEV_FECHA ) TAG HIST_NO
CREATE TABLE IMEDI (DOCU_NO C(12),DOCU_FECHA D,TIDO_COD C(6),HIST_NO C(12),PACI_DES C(60),;
   PACI_EDAD C(20),FEADMISION D,FEALTA D,DOCU_DES M)

nValRet = SQLEXEC( nConex, "select * from datosp order by 1", "ttx")

SCAN
   num = PADL( ALLT( STRT( TTX.N_HISTO, "+" ) ), 12 )
   num = STRT( num, "-" )
   fec = IIF( TTOD( TTX.FECHA_I ) > dini, TTOD( TTX.FECHA_I ), CTOD( "" ) )
   sex = IIF( TTX.SEXO = "F", "FEMENINO", "MASCULINO" )
   pes = IIF( !ISNULL( TTX.PESOAN ), VAL( STRT( TTX.PESOAN, ".", "," ) ), 0 )
   fla = IIF( !ISNULL( TTX.FORMULA ), ALLT( TTX.FORMULA ), "" )
   INSERT INTO PACIE (HIST_NO,HIST_FECHA,PACI_TIPO,PACI_APE,PACI_NOM,PACI_LUNAC,;
                      PACI_FENAC,PACI_SEXO,PACI_PESO,PACI_FORMU,REFERIDO,PACI_MADRE,;
                      PACI_PADRE,PACI_DIREC,PACI_TELFS,HIST_OANTE,HIST_OVACU) VALUES ;
                     (num,fec,UPPER(TTX.TIPO_P),ALLT(TTX.APELLIDOS),ALLT(TTX.NOMBRES),;
                      ALLT(TTX.LUGAR_N),TTOD(TTX.FECHA_N),sex,pes,fla,TTX.REFERIDOP,TTX.MADRE,;
                      TTX.PADRE,ALLT(TTX.DIREC),ALLT(TTX.TELF),ALLT(TTX.ANTECED),ALLT(TTX.VACUNAS))
   IF ( !EMPTY( TTX.HISTORIA ) )
      lon = MEMLINES( TTX.HISTORIA )
      cad = ""
      fec = CTOD( "" )
      FOR i = 1 TO lon
         lin = ALLT( MLINE( TTX.HISTORIA, i ) )
         IF ( ISDIGIT( lin ) AND AT( "-", lin ) = 3 )
            cad2 = IIF( !EMPTY( fec ), SUBS( cad, 12 ), cad )
            INSERT INTO HISTO (HIST_NO,DIEV_FECHA,DIEV_DES) VALUES (num,fec,cad2)
            cad = lin
            fec = CTOD( SUBS( lin, 1, 2 )+"/"+SUBS( lin, 4, 2 )+"/"+SUBS( lin, 7, 4 ))
            fec = IIF( fec > dini, fec, CTOD( "" ) )
         ELSE
            cad = (cad+lin)
         ENDIF
         cad = STRTRAN( cad, "   " )
      NEXT
      cad2 = IIF( !EMPTY( fec ), SUBS( cad, 12 ), cad )
      INSERT INTO HISTO (HIST_NO,DIEV_FECHA,DIEV_DES) VALUES (num,fec,cad2)
   ENDIF
ENDSCAN

DELETE FROM HISTO WHERE EMPTY( DIEV_DES )

i = 100
nValRet = SQLEXEC( nConex, "select * from infmedcea", "ttx")
SQLDISCONNECT( nConex )

SCAN
   num  = PADL( ALLT( STRT( TTX.N_HISTO, "+" ) ), 12 )
   num  = STRT( num, "-" )
   fec  = IIF( TTOD( TTX.FECHAI ) > dini, TTOD( TTX.FECHAI ), CTOD( "" ) )
   cad  = STR( i, 12 )
   cad2 = (ALLT( TTX.APELLIDOS )+" "+ALLT( TTX.NOMBRES ))
   fla = IIF( !ISNULL( TTX.FECHAA ), TTOD( TTX.FECHAA ), CTOD( "" ) )
   fla = IIF( fla > dini, fla, CTOD( "" ) )
   INSERT INTO IMEDI (DOCU_NO,DOCU_FECHA,TIDO_COD,HIST_NO,PACI_DES,PACI_EDAD,FEADMISION,FEALTA,DOCU_DES) VALUES ;
                     (cad,fec,"003",num,cad2,"...",fec,fla,ALLT( TTX.INFO ))
   i = i + 1
ENDSCAN

nOpc = MESSAGEBOX( "¿ ESTA SEGURO ?", 4 + 32 + 256, "Conforme" )
IF ( nOpc = 6 )
   USE (cRutadbf+"historia.dbf")  IN 0 EXCL
   USE (cRutadbf+"dievoluc.dbf")  IN 0 EXCL
   USE (cRutadbf+"documtos.dbf")  IN 0 EXCL
   USE (cRutadbf+"correlat.dbf")  IN 0 EXCL
   USE (cRutadbf+"cvacunas.dbf")  IN 0 EXCL
   SELECT HISTORIA
   ZAP
   APPEND FROM ("PACIE")
   SELECT DIEVOLUC
   ZAP
   APPEND FROM ("HISTO")
   SELECT DOCUMTOS
   ZAP
   APPEND FROM ("IMEDI")
   UPDATE CORRELAT SET CORR_NUM = 6000 WHERE CORR_COD = "HISTORIA"
   UPDATE CORRELAT SET CORR_NUM = (VAL( DOCUMTOS.DOCU_NO ) + 100) WHERE CORR_COD = "DOCUMENTO"
   ZAP IN CVACUNAS
   MESSAGEBOX( "DATOS FUERON COPIADOS", 64, "Mensaje" )
ENDIF

CLOSE DATABASES ALL

Reindexa( )