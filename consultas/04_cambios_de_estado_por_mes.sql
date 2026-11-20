-- 04 · Tabla de referencias cruzadas: cambios de estado por mes
-- Filas = mes, columnas = estado. Access pide la fecha inicial y la final.
PARAMETERS [Fecha inicial] DateTime, [Fecha final] DateTime;
TRANSFORM Count(pv.PuntoID) AS Cantidad
SELECT Format(pv.FechaEstado, "yyyy-mm") AS Mes
FROM PuntosVenta AS pv
WHERE pv.FechaEstado Between [Fecha inicial] And [Fecha final]
GROUP BY Format(pv.FechaEstado, "yyyy-mm")
PIVOT pv.Estado In ("ACTIVO", "PENDIENTE", "SUSPENDIDO", "CERRADO");
