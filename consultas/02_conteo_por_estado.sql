-- 02 · Cuántos puntos hay en cada estado y qué porcentaje representan
-- DCount (función de dominio de Access) da el total de la tabla para calcular el %.
SELECT
    pv.Estado,
    Count(*) AS Cantidad,
    Format(Count(*) / DCount("*", "PuntosVenta"), "0.0%") AS Porcentaje
FROM PuntosVenta AS pv
GROUP BY pv.Estado
ORDER BY Count(*) DESC;
