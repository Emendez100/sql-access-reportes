-- 03 · Calidad de datos: puntos que NO tienen coordenadas
-- Estos puntos desaparecen de la consulta 01 porque usa INNER JOIN.
-- Con LEFT JOIN + "Is Null" se encuentran los que faltan.
SELECT
    pv.PuntoID AS [ID Punto],
    d.NombreComercial AS [Nombre Comercial],
    d.Red,
    d.Provincia,
    pv.Estado
FROM (PuntosVenta AS pv
    INNER JOIN Direcciones AS d ON pv.PuntoID = d.PuntoID)
    LEFT JOIN Geolocalizacion AS g ON pv.PuntoID = g.PuntoID
WHERE g.PuntoID Is Null
ORDER BY d.Red, pv.PuntoID;
