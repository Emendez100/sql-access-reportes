-- 05 · Resumen por red y provincia: total, activos, cerrados y % de activos
SELECT
    d.Red,
    d.Provincia,
    Count(*) AS Total,
    Sum(IIf(pv.Estado = "ACTIVO", 1, 0)) AS Activos,
    Sum(IIf(pv.Estado = "CERRADO", 1, 0)) AS Cerrados,
    Format(Sum(IIf(pv.Estado = "ACTIVO", 1, 0)) / Count(*), "0%") AS [% Activos]
FROM PuntosVenta AS pv
    INNER JOIN Direcciones AS d ON pv.PuntoID = d.PuntoID
GROUP BY d.Red, d.Provincia
ORDER BY d.Red, Count(*) DESC;
