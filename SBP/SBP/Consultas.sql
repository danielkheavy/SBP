--USE RESTAURANT6 
SELECT * FROM tipo where tipo='1'
SELECT * FROM tipo where tipo='2'
--SELECT * FROM " & gocabeza & "   where  local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic
SELECT * FROM factura where  local='01' and tipo='1' and serie='B01' and numero='57'
SELECT * FROM factura
SELECT * FROM tipo where  tipo='1'
SELECT * FROM PRODUCTO
UPDATE producto SET PUERTOIMPRESION='Snagit 9', PUERTOIMPRESION1='Snagit 9', PUERTOIMPRESION2='Snagit 9',PUERTOIMPRESION3='Snagit 9' 

-- consulta de productos con stock minimo----
select P.PRODUCTO,P.DESCRIPCIO,P.UNIDAD,P.FACTOR,P.MINIMO,A.SALDO AS SALDO_ACTUAL, P.COSTOU, P.COSTOP, A.SALDO * P.COSTOU AS TOTAL, P.FAMILIA, P.SUBFAMILIA   
from producto AS P, ALMACEN AS A 
WHERE P.PRODUCTO=A.PRODUCTO AND A.SALDO < P.MINIMO
order by P.MINIMO DESC

select * from producto where producto like '%' order by familia,Subfamilia,descripcio
Select * from almacen where local='01' and  producto='13' and bodega='01'
SELECT * FROM producto  
select * from almacen

SELECT * FROM precios where producto='132'  minimo11
