Select T1.idstrNoPoliza
 ,T2.strIdAseguradora
 ,T2.strNomAse
 ,T1.strNoReciboInt
 ,T2.CardCode
 ,LEFT(T2.CardName,30) AS CardName
 ,T6.strRFC AS strRFC
 ,LEN(T6.strRFC) AS LargoRFC
 ,T6.bintIdCteOperacion
 ,T6.bintIdCliente
 ,T2.intIva
 ,T5.strCveSap AS Asociado
 ,T3.strCveSap AS Ejecutivo
 ,T5.intMTipoContacto as intMTipoContacto1
 ,T3.intMTipoContacto AS intMTipoContacto
 ,T1.dteAplicacion
 ,T2.strIdProd AS ItemCode
 ,T2.strNomProd AS ItemName
 ,T1.dcmPrimaR + T1.dcmDP + T1.dcmRPF AS PrecioNet
 ,T4.strObservaciones
 ,T2.intComR
 FROM appOrdenTrabajo T0											INNER 
 JOIN appEmisionPolizaRecibosS T1									INNER 
 JOIN appEmisionPolizas T2 ON T1.bintIdPoliza = T2.bintIdPoliza 
             ON T0.strIdPoliza = T2.idstrNoPoliza	INNER 
 JOIN appMtrContactoAPIASSA T3										INNER 
 JOIN appOPVenta T4 ON T3.SlpCode = T4.SlpCode 
      ON T0.strNOp = T4.strNOp						LEFT OUTER 
 JOIN appMtrContactoAPIASSA T5 ON T4.ChnCrdCode = T5.SlpCode
 JOIN appMtrClientes T6 ON T6.CardCode = T4.CardCode
 WHERE T1.bintIdStatusRe = 7
   AND T1.BitSap = 0 
 ORDER BY T1.idstrNoPoliza