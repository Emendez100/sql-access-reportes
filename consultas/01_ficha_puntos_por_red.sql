-- 01 · Ficha completa de los puntos de una red (consulta de creación de tabla)
-- Une las 4 tablas y guarda el resultado en [Tbl_Ficha de Puntos] para usarlo en reportes.
-- Al ejecutarla, Access pide el código de red (por ejemplo RN01).
PARAMETERS [Red a consultar] Text ( 255 );
SELECT
    pv.PuntoID AS [ID Punto],
    m.MarcaID AS [ID Marca],
    pv.Estado,
    pv.FechaEstado AS [Fecha de Estado],
    pv.CodigoClase AS [Codigo de Clase],
    d.NombreComercial AS [Nombre Comercial],
    pv.RazonSocial AS [Razon Social],
    d.Red,
    d.Subred,
    d.Direccion1 AS [Direccion 1],
    d.Direccion2 AS [Direccion 2],
    d.Ciudad,
    d.Provincia,
    d.CodigoPostal AS [Codigo Postal],
    g.Longitud,
    g.Latitud,
    d.Telefono1 AS [Telefono 1],
    d.Telefono2 AS [Telefono 2],
    pv.Comentarios,
    pv.PuntoIDAnterior AS [ID Punto Anterior],
    d.TipoPunto AS [Tipo de Punto],
    g.Canal
INTO [Tbl_Ficha de Puntos]
FROM ((Direcciones AS d
    INNER JOIN PuntosVenta AS pv ON d.PuntoID = pv.PuntoID)
    INNER JOIN Marcas AS m ON pv.PuntoID = m.PuntoID)
    INNER JOIN Geolocalizacion AS g ON m.PuntoID = g.PuntoID
WHERE d.Red = [Red a consultar];
