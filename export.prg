*#DEFINE testing

#DEFINE MSG_ERR 16
#DEFINE MSG_WAR 48
#DEFINE MSG_INFO 64
#DEFINE CRLF CHR(13) + CHR(10)

LPARAMETERS sCurrentMonth, lGetJornal

SET TEXTMERGE ON
SET ENGINEBEHAVIOR 70

LOCAL sDB

sDB = ".\RECIBOS.DBC"
sPathDBSICOSS = SYS(5) + SYS(2003) + "\sijp12\sijp12.mdb"
sPathFieldInfo = ".\fieldinfo.txt"
sPathPlanilla = SYS(5) + SYS(2003) + "\Recibo.xls"
sOutFilename = ".\eSICOSS.txt"

#IFDEF testing
	sCurrentMonth = "10"
	lGetJornal = .F.
#ENDIF

lGetJornal = (lGetJornal == "T")

nDefaultJornada = IIF(lGetJornal, 2.5, 1)

OPEN DATABASE (sDB) SHARED NOUPDATE
USE recibo AGAIN IN 0

TEXT TO sSQLStr NOSHOW
DRIVER=Microsoft Access Driver (*.mdb);DBQ=<<sPathDBSICOSS>>;PWD=naDdePraKciN
ENDTEXT

? "Accediendo a SICOSS..."

xSQLConn = SQLSTRINGCONNECT(sSQLStr)
sPeriod = "'" + ALLTRIM(STR(YEAR(DATE()))) + PADL(sCurrentMonth, 2, '0') + "'"
sQuery = "SELECT * FROM 22CUILes WHERE [Período]=" + sPeriod

IF xSQLConn >= 1
	IF ! SQLEXEC(xSQLConn, sQuery, "C_SICOSS") == 1
		MESSAGEBOX("No se pudo acceder a la DB de SICOSS. " ;
					+ "Saliendo...", MSG_ERR, "Error", -1)
		#IFNDEF testing
			QUIT
		#ELSE
			CANCEL
		#ENDIF
	ENDIF
ENDIF

SELECT * FROM recibo ;
	WHERE ALLTRIM(STR(MONTH(feclia)))==sCurrentMonth ;
	AND YEAR(feclia)==YEAR(DATE()) ;
	ORDER BY cuil ASC ;
	INTO CURSOR C_Rec READWRITE


* Compatibilizo con SICOSS
UPDATE C_Rec ;
	SET cuil = STRTRAN(cuil, "-", "") ;
	WHERE .T.

* Agrego jornada
SELECT *, (nDefaultJornada) AS jornada ;
	FROM C_Rec ;
	INTO CURSOR C_Rec READWRITE

IF lGetJornal
	InsertJornal(sPathPlanilla)
ENDIF

SELECT r.*, s.* FROM C_Rec r ;
	INNER JOIN C_SICOSS s ;
	ON r.cuil == s.cuil ;
	INTO CURSOR C_Def READWRITE

sFieldInfo = FILETOSTR(sPathFieldInfo)
nFieldCount = OCCURS(CRLF, sFieldInfo)

DIMENSION aFieldSpecs(nFieldCount, 5) && name | type | length | decimal | pad
ALINES(aFlds, sFieldInfo, 4)

* tipo | nombre | length | formula
* m|apo_adi_ooss|7.2|ICASE(ooss == 126205, 100, 0)
sSQLQuery = ""
i = 1
FOR EACH sFld IN aFlds
	sType = GETWORDNUM(sFld, 1, "|")
	sName = GETWORDNUM(sFld, 2, "|")
	sLen = GETWORDNUM(sFld, 3, "|")
	sFormula = GETWORDNUM(sFld, 4, "|")

	lPad = (SUBSTR(sLen, 1, 1) == "0")
	nLen = INT(VAL(sLen))
	nDec = INT(MOD(VAL(sLen), 1) * 10)

	aFieldSpecs(i, 1) = sName
	aFieldSpecs(i, 2) = sType
	aFieldSpecs(i, 3) = nLen
	aFieldSpecs(i, 4) = nDec
	aFieldSpecs(i, 5) = lPad

	sSQLQuery = sSQLQuery + " (&sFormula) AS &sName,"
	i = i + 1
NEXT

* Elimino la última ','
sSQLQuery = LEFT(sSQLQuery, LEN(sSQLQuery) - 1)

? sSQLQuery

SELECT &sSQLQuery ;
	FROM C_Def ;
	INTO CURSOR C_Def ;
	GROUP BY cuil

nEmplCount = _TALLY

* Escribo el txt
STRTOFILE("", sOutFilename, 0)
SCAN
	FOR i = 1 TO nFieldCount
		sVal = FormatAsString(&aFieldSpecs(i, 1), aFieldSpecs(i, 2), ;
				aFieldSpecs(i, 3), aFieldSpecs(i, 4), aFieldSpecs(i, 5))

		STRTOFILE(sVal, sOutFilename, 1)
	ENDFOR

	STRTOFILE(CRLF, sOutFilename, 1)
ENDSCAN

MESSAGEBOX("Todo listo. Se procesaron " + ALLTRIM(STR(nEmplCount)) ;
			+ " empleados.", MSG_INFO, "ExportSICOSS", -1)

#IFNDEF testing
	QUIT
#ENDIF


**

FUNCTION InsertJornal(sPathPlanilla)
	LOCAL oXL, oRange, nCol, sCUIL, nJornal

*	oXL = GETOBJECT(sPathPlanilla)
	oXL = CREATEOBJECT("Excel.Application")
	oXL.Application.Workbooks.Open(sPathPlanilla)

	IF VARTYPE(oXL) == "O"
		oXL.Visible = .F.
		oXL.DisplayAlerts = .F.

		oXL.Worksheets("Empleados").Unprotect()

		oRange = oXL.Worksheets("Empleados").Cells(1,1).Currentregion

		nCol = 4

		SELECT C_Rec

		*sCUIL = STRTRAN(oRange.Cells(6, nCol).Value, "-", "")
		sCUIL = "CUIT EJEMPLO"
		DO WHILE ! sCUIL == ""
			IF nCol == 1000 && Proteccion de loop infinito
				EXIT
			ENDIF

			sCUIL = STRTRAN(oRange.Cells(6, nCol).Value, "-", "")
			nJornal = oRange.Cells(27, nCol).Value
			UPDATE C_Rec SET jornada = nJornal WHERE cuil == sCUIL
			nCol = nCol + 1
		ENDDO

		oXL.Workbooks.Close()

		RELEASE oXL
	ENDIF
ENDFUNC

FUNCTION FormatAsString(xVal, sType, nLen, nDec, lPad)
	sRet = ""

	DO CASE
	CASE sType == "n" OR sType == "m"
		sRet = STR(xVal, nLen + 1, nDec)
		sRet = STRTRAN(sRet, ".", ",")
		sRet = PADL(ALLTRIM(sRet), nLen + nDec, IIF(lPad, "0", " "))
	CASE sType == "l"
		sRet = IIF(xVal, "T", "F")
	OTHERWISE
		sRet = SUBSTR(xVal, 1, nLen)
	ENDCASE

	RETURN sRet
ENDFUNC