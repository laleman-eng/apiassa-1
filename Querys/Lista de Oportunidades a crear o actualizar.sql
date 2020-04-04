Select T0.bintIDOpVen
 ,T1.CardCode
 ,RTRIM(LEFT(T1.CardName,100)) AS CardName
 ,T1.bintIdCteOperacion
 ,T1.strRFC AS strRFC
 ,LEN(T1.strRFC) AS LargoRFC
 ,T1.bintIdCliente
 ,CAST(T3.strCveSap AS VarChar(20)) AS Asociado
 ,CAST(T2.strCveSap AS VarChar(20)) AS Ejecutivo
 ,CAST(A1.strCveSap AS VarChar(20)) AS AsociadoB
 ,CAST(A0.strCveSap AS VarChar(20)) AS EjecutivoB
 ,T0.strObservaciones
 ,T0.strTitular
 ,T0.intTPla
 ,T0.dteInicio
 ,T0.dteCierreEstimado
 ,T0.dcmMToP
 ,T0.IntRate
 ,T0.IntId
 ,T0.ChnCrdCode
 ,T0.strNOp 
 ,T0.bigIntEdo
 ,T0.intPerTPla 
 FROM appOPVenta T0
 JOIN appMtrClientes T1 ON T1.CardCode = T0.CardCode		LEFT
 JOIN appMtrContactoAPIASSA T2	ON T2.SlpCode = T0.SlpCode	LEFT
 JOIN appMtrContactoAPIASSA T3 ON T3.SlpCode = T0.ChnCrdCode	LEFT
 JOIN appMtrContactoAPIASSA A0	ON A0.SlpCode = T0.SlpCodeB		LEFT
 JOIN appMtrContactoAPIASSA A1	ON A1.SlpCode = T0.ChnCrdCodeB
 JOIN appOPVentaDet T4 ON T4.strNOp = T0.strNOp
 WHERE(Not (T0.bintIDOpVen Is NULL))
 AND (T4.bitSap = 0)
 -- AND T0.strNOp = 'OV-A08GMG32010-001' 'DEJAR COMO COMENTARIO ES SOLO PARA PRUEBA
 GROUP BY T0.bintIDOpVen
 ,T1.CardCode
 ,RTRIM(LEFT(T1.CardName,100))
 ,T1.bintIdCteOperacion
 ,T1.strRFC
 ,T1.bintIdCliente
 ,CAST(T3.strCveSap AS VarChar(20))
 ,CAST(T2.strCveSap AS VarChar(20))
 ,T0.strObservaciones
 ,T0.strTitular
 ,T0.intTPla
 ,T0.dteInicio
 ,T0.dteCierreEstimado
 ,T0.dcmMToP
 ,T0.IntRate
 ,T0.IntId
 ,T0.ChnCrdCode
 ,T0.strNOp 
 ,T0.bigIntEdo
 ,T0.intPerTPla
 ,CAST(A1.strCveSap AS VarChar(20))
 ,CAST(A0.strCveSap AS VarChar(20))
