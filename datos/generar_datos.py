"""Genera los datos FICTICIOS de la base de ejemplo (puntos de venta).

Todos los nombres, teléfonos y direcciones son inventados.
Uso:  python generar_datos.py
Crea 4 archivos CSV en esta misma carpeta.
"""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

random.seed(2026)
CARPETA = Path(__file__).parent
N = 400

PROVINCIAS = {
    "SAN JOSE": (["SAN JOSE", "ESCAZU", "DESAMPARADOS", "CIUDAD COLON", "SANTA ANA"], "10"),
    "ALAJUELA": (["ALAJUELA", "GRECIA", "SAN RAMON", "SAN CARLOS"], "20"),
    "CARTAGO": (["CARTAGO", "PARAISO", "TURRIALBA"], "30"),
    "HEREDIA": (["HEREDIA", "BARVA", "SANTO DOMINGO"], "40"),
    "GUANACASTE": (["LIBERIA", "NICOYA", "SANTA CRUZ"], "50"),
    "PUNTARENAS": (["PUNTARENAS", "QUEPOS", "PEREZ ZELEDON"], "60"),
    "LIMON": (["LIMON", "GUAPILES", "SIQUIRRES"], "70"),
}
# Centro aproximado de cada provincia para generar coordenadas
CENTROS = {
    "SAN JOSE": (9.93, -84.08), "ALAJUELA": (10.02, -84.21), "CARTAGO": (9.86, -83.92),
    "HEREDIA": (10.00, -84.12), "GUANACASTE": (10.63, -85.44), "PUNTARENAS": (9.98, -84.83),
    "LIMON": (9.99, -83.03),
}
TIPOS_NEGOCIO = ["PULPERIA", "FARMACIA", "SUPER", "FERRETERIA", "LIBRERIA", "MINISUPER", "BAZAR"]
NOMBRES = ["LA ESQUINA", "EL BUEN PRECIO", "SAN MARTIN", "LOS ROBLES", "LA ECONOMIA", "EL PROGRESO",
           "LA FAMILIA", "DON PEPE", "LAS BRISAS", "EL CENTRO", "LA AMISTAD", "SANTA MARIA"]
APELLIDOS = ["MORA", "ROJAS", "VARGAS", "JIMENEZ", "SOLIS", "CASTRO", "ARAYA", "CHAVES", "QUESADA", "HERRERA"]
SOCIEDAD = ["S.A.", "S.R.L.", "LTDA."]
ESTADOS = ["ACTIVO"] * 70 + ["SUSPENDIDO"] * 12 + ["CERRADO"] * 13 + ["PENDIENTE"] * 5
CLASES = ["A1", "A2", "B1", "B2", "C1"]
REDES = {"RN01": ["N1", "N2"], "RN02": ["E1", "O1"], "RN03": ["C1"], "RS01": ["S1", "S2"]}
COMENTARIOS = {
    "ACTIVO": ["", "", "", "ALTA NUEVA", "REACTIVADO"],
    "SUSPENDIDO": ["FALTA DOCUMENTO", "REVISION DE CUMPLIMIENTO", "SIN MOVIMIENTO 90 DIAS"],
    "CERRADO": ["CIERRE A SOLICITUD DEL CLIENTE", "CAMBIO DE DUENO", "NEGOCIO CERRADO"],
    "PENDIENTE": ["EN TRAMITE", "ESPERANDO CAPACITACION"],
}
MARCAS = ["M01", "M02", "M03"]
CANALES = ["POS", "WEB", "APP"]


def telefono(fijo=True):
    return f"{random.choice('24') if fijo else random.choice('678')}{random.randint(0, 999):03d}-{random.randint(0, 9999):04d}"


def main():
    inicio = date(2023, 1, 1)
    puntos, direcciones, marcas, geo = [], [], [], []
    telefonos_usados = []
    for i in range(N):
        pid = 500100 + i * 7
        estado = random.choice(ESTADOS)
        fecha = inicio + timedelta(days=random.randint(0, 1400))
        prov = random.choice(list(PROVINCIAS))
        ciudades, cp = PROVINCIAS[prov]
        red = random.choice(list(REDES))
        nombre = f"{random.choice(TIPOS_NEGOCIO)} {random.choice(NOMBRES)}"
        anterior = str(pid - 300000) if random.random() < 0.08 else ""
        puntos.append({
            "PuntoID": pid, "Estado": estado, "FechaEstado": fecha.isoformat(),
            "CodigoClase": random.choice(CLASES),
            "RazonSocial": f"INVERSIONES {random.choice(APELLIDOS)} Y {random.choice(APELLIDOS)} {random.choice(SOCIEDAD)}",
            "Comentarios": random.choice(COMENTARIOS[estado]), "PuntoIDAnterior": anterior,
        })
        # ~4 % de teléfonos repetidos a propósito, para la consulta de calidad de datos
        tel1 = random.choice(telefonos_usados) if telefonos_usados and random.random() < 0.04 else telefono()
        telefonos_usados.append(tel1)
        direcciones.append({
            "PuntoID": pid, "NombreComercial": nombre, "Red": red, "Subred": random.choice(REDES[red]),
            "Direccion1": f"{random.randint(25, 400)} M {random.choice(['NORTE', 'SUR', 'ESTE', 'OESTE'])} DE LA IGLESIA",
            "Direccion2": random.choice(["", "", "LOCAL 2", "FRENTE AL PARQUE", "CONTIGUO AL BANCO"]),
            "Ciudad": random.choice(ciudades), "Provincia": prov,
            "CodigoPostal": f"{cp}{random.randint(101, 999)}",
            "Telefono1": tel1, "Telefono2": telefono(fijo=False) if random.random() < 0.6 else "",
            "TipoPunto": random.choice(["PROPIO", "AFILIADO", "AFILIADO"]),
        })
        marcas.append({"PuntoID": pid, "MarcaID": random.choice(MARCAS)})
        # ~6 % sin geolocalización, para la consulta de datos faltantes
        if random.random() > 0.06:
            lat, lon = CENTROS[prov]
            geo.append({"PuntoID": pid, "Latitud": round(lat + random.uniform(-0.15, 0.15), 6),
                        "Longitud": round(lon + random.uniform(-0.15, 0.15), 6),
                        "Canal": random.choice(CANALES)})

    for nombre, filas in [("puntos_venta", puntos), ("direcciones", direcciones),
                          ("marcas", marcas), ("geolocalizacion", geo)]:
        with open(CARPETA / f"{nombre}.csv", "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=list(filas[0]))
            w.writeheader()
            w.writerows(filas)
        print(f"{nombre}.csv: {len(filas)} filas")


if __name__ == "__main__":
    main()
