-- 06 · Calidad de datos: el mismo teléfono registrado en más de un punto
-- La subconsulta encuentra los teléfonos repetidos; la consulta principal muestra a quién pertenecen.
SELECT
    d.Telefono1 AS Telefono,
    d.PuntoID AS [ID Punto],
    d.NombreComercial AS [Nombre Comercial],
    d.Ciudad
FROM Direcciones AS d
WHERE d.Telefono1 In (
    SELECT Telefono1
    FROM Direcciones
    GROUP BY Telefono1
    HAVING Count(*) > 1
)
ORDER BY d.Telefono1, d.PuntoID;
