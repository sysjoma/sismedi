* SYSJOMA
* Conversión de archivo de access HISTORIA.MDB
* 11/06/2009

LOCAL cCadConex, nConex, nValRet, i, lon, cad, lin, fec, cad2, num, sex, pes, fla, dini

cCadConex = "DRIVER={Microsoft Access Driver (*.mdb)};" + ;
            "DBQ=c:\historia\historiavb.mdb;"
dini = CTOD( "01/01/1930" )

SQLSETPROP( 0, "DispLogin" , 3 )

nConex = SQLSTRINGCONNECT( cCadConex )

IF ( nConex <= 0 )
   AERROR( laErr )
   MESSAGEBOX( "No se pudo conectar a SQL Server. Error: 1"+CHR( 13 )+laErr[ 2 ] )
   RETURN
ENDIF

CREATE CURSOR PACIE (HIST_NO C(12),HIST_FECHA D,PACI_TIPO C(30),PACI_APE C(40),PACI_NOM C(40),;
   PACI_LUNAC C(40),PACI_FENAC D,PACI_SEXO C(9),PACI_PESO N(8,3),PACI_FORMU C(30),REFERIDO C(40),;
   PACI_MADRE C(40),PACI_PADRE C(40),PACI_DIREC C(90),PACI_TELFS C(40),HIST_OANTE C(60),HIST_OVACU C(60))
INDEX ON HIST_NO TAG HIST_NO
CREATE CURSOR HISTO (HIST_NO C(12),DIEV_FECHA D,DIEV_DES C(200))
INDEX ON HIST_NO+DTOS( DIEV_FECHA ) TAG HIST_NO
CREATE CURSOR IMEDI (DOCU_NO C(12),DOCU_FECHA D,TIDO_COD C(6),HIST_NO C(12),PACI_DES C(60),;
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

SELECT PACIE
COPY TO BBB
SELECT HISTO
COPY TO CCC
SELECT IMEDI
COPY TO DDD

CLOSE DATABASES ALL