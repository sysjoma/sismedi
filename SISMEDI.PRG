PUBLIC VCOPY, UCOPY, BLOER, PCID, FHIN, FHFI, NUSER, ROLOC2, ;
       cRutadbf, nUacce, cPrn, aPrn( 1 ), cNoEmpre, cTitRep1, xVar
LOCAL lPase

STORE "" TO VCOPY, UCOPY, BLOER, PCID, cPrn, cTitRep1, xVar
cRutadbf = (SYS( 2003 )+"\dbf\")
NUSER    = SPAC( 12 )
nUacce   = 3
cNoEmpre = "Dra. Andry J. Moreno Alvarado"

_SCREEN.VISIBLE     = .F.
_SCREEN.BACKCOLOR   = RGB( 255, 255, 255 )
_SCREEN.CAPTION     = "SISMEDI 3.3  /  Elaborado por: José Alemán  0424-849.24.55  0281-276.21.51  josedevene@hotmail.com"
_SCREEN.WINDOWSTATE = 2
_SCREEN.MAXBUTTON   = .F.

SET TALK OFF
SET SAFETY OFF
SET CONFIRM OFF
SET NOTIFY OFF
SET DATE TO SHORT
SET SYSFORMATS ON
SET CLOCK STATUS
SET DELETED ON
SET EXCLUSIVE OFF
SET REPROCESS TO 5 SECONDS
SET SYSMENU OFF
CLEAR
CAPSLOCK( .T. )
*NUMLOCK( .T. )

SET PATH TO FRX, IMG
SET PROCEDURE TO FSYSJOMA

IF ( LeerConf( "3sismedi.ini" ) )
   =VeriSist( )
   _SCREEN.VISIBLE = .T.
   DO FORM ACCESO TO lPase
   IF ( lPase )
      =Escratex( "hlogins.txt", (PCID+" en "+TTOC( DATETIME( ) )+" "+LOWER( NUSER )) )
      _SCREEN.ADDOBJECT( "Image1", "image" )
      _SCREEN.Image1.PICTURE = "f0.jpg"
      _SCREEN.Image1.TOP     = _SCREEN.HEIGHT / 2 - _SCREEN.Image1.HEIGHT / 2
      _SCREEN.Image1.LEFT    = _SCREEN.WIDTH / 2 - _SCREEN.Image1.WIDTH / 2
      _SCREEN.Image1.STRETCH = 1
      _SCREEN.Image1.VISIBLE = .T.
      oToolBar = NEWOBJECT( "mitoolbar", "mitoolbar.vcx" )
      oToolBar.SHOW
      DO MENUPRIN.MPR
      READ EVENTS
      =Escratex( "hlogins.txt", (PCID+" sa "+TTOC( DATETIME( ) )+" "+LOWER( NUSER )) )
   ENDIF
ENDIF

RETURN




FUNCTION VeriSist( )
  LOCAL cFile

  cFile = "ERR00.BAK"

  IF ( FILE( cFile ) )
     WAIT WINDOWS "FATAL ERROR IN STRUCT EXE/DLL/DBF"
     QUIT
  ENDIF

  IF ( !EMPTY( BLOER ) AND DTOS( DATE( ) ) >= ("20"+BLOER) )
     DELETE FILE *.CDX
     =FCREATE( cFile )
  ENDIF

  RETURN
ENDFUNC




FUNCTION AbrirXLS( cFileori, cFilecop )
   LOCAL oExcel

   cFileori = (SYS( 5 )+SYS( 2003 )+"\doc\"+cFileori+".xls")
   cFilecop = (SYS( 5 )+SYS( 2003 )+"\doc\"+cFilecop+".xls")

   COPY FILE (cFileori) TO (cFilecop)
   oExcel = CREATEOBJECT( "Excel.Application" )
   oExcel.Workbooks.Open( cFilecop )
   oExcel.Visible=.T.

   RETURN
ENDFUNC