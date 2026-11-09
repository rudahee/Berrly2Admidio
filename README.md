# berrly-to-admidio

Script para migrar el listado de afiliados de [Berrly](https://berrly.app/)
a [Admidio](https://www.admidio.org/): convierte la exportación en Excel de
Berrly (`data.xlsx`) en un CSV (`data.csv`) importable directamente en
Admidio.

## Qué hace

Lee `data.xlsx` (exportación de Berrly) y genera `data.csv` con las columnas
que espera el importador de Admidio:

```
Apellido, Nombre, Calle, Código Postal, Ciudad, Pais,
Teléfono, Teléfono 2, WhatsApp, E-Mail,
Numero de afiliado, Fecha de Afiliacion, Tipo de Documento,
Documento de Identidad, Genero, Fecha de Nacimiento, Notas,
¿Permite el almacenamiento y tratamiento de datos personales?
```

Durante la conversión:

- Limpia espacios y comillas sueltas de los campos de texto.
- Normaliza teléfonos (quita el `.0` de pandas, el prefijo `+34` y los espacios).
- Combina varios campos de teléfono/email de Berrly en uno solo.
- Formatea las fechas como `d.m.aaaa`, el formato que espera Admidio.

## Uso

```bash
pip install pandas openpyxl
python main.py
```

El script espera encontrar `data.xlsx` (la exportación de Berrly) en el
mismo directorio y escribe `data.csv` junto a él, listo para importar en
Admidio.
