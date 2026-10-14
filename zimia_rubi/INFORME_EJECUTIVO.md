# Zimia V1: establecimientos de Rubí. Informe ejecutivo

**Fecha de la investigación:** 2026-10-14 · **Cartografía base:** Overture Maps 2026-09-23.0 · **Municipio:** Rubí (Barcelona)

Entregables:
- `Zimia_Establecimientos_Rubi_V1.xlsx`: hojas ESTABLECIMIENTOS, PENDIENTES DE VERIFICACIÓN, RESUMEN, METODOLOGÍA y una hoja auxiliar, DESCARTADOS, con los candidatos excluidos y los duplicados fusionados.
- `Zimia_Establecimientos_Rubi_V1.csv`: las 31 columnas V1 en UTF-8, con los NULL exportados como celdas vacías.

## 1. Cuántos establecimientos se han encontrado

**79 registros**: 76 activos o pendientes de revisión y 3 cerrados confirmados. Los cerrados se conservan para que no vuelvan a entrar desde directorios antiguos.

Además se descartaron 20 candidatos (fuera del término municipal, fuera de alcance o sin confirmar) y se fusionaron 14 duplicados.

No se puede afirmar que estén todos los establecimientos de Rubí. Ver el apartado 5.

## 2. Categorías con más presencia (76 registros no cerrados)

| Categoría | n.º | % |
|---|---|---|
| supermercado | 40 | 52,6 % |
| tienda_mascotas | 12 | 15,8 % |
| veterinario | 8 | 10,5 % |
| otros (peluquerías caninas o felinas y servicios para animales) | 7 | 9,2 % |
| tienda_natural | 5 | 6,6 % |
| agrotienda | 2 | 2,6 % |
| garden | 1 | 1,3 % |
| hipermercado | 1 | 1,3 % (Esclat, 3.058 m²) |

- **Tipo de empresa:** 39 de cadena, 33 independientes y 4 sin determinar.
- **Supermercados por cadena:** Consum/Charter 7, Condis 9 (3 pendientes), Dia 4, Mercadona 3, Lidl 2, Caprabo 2, bonÀrea 2, y ALDI, Alcampo, Sorli, Carrefour Express, SUMA y Covirán con 1 cada una.

No hay en Rubí ninguna tienda de las grandes cadenas especializadas (Kiwoko, Tiendanimal). La "Tiendanimal Rubí" de los directorios está en Sant Cugat.

## 3. Porcentaje verificado

| Nivel | n.º | % |
|---|---|---|
| confirmado | 3 | 3,8 % |
| parcialmente_confirmado | 56 | 70,9 % |
| no_verificado | 20 | 25,3 % |

- **Identidad y ubicación verificadas** (confirmado + parcial): 59 de 79, un **74,7 %**.
- **Venta de alimentación confirmada** (`si`): 14 establecimientos, de los cuales 11 activos y 3 pendientes de revisión.
- **Perros confirmado:** 3. **Gatos confirmado:** 3. Son Sweet Doggy, Petslovers Rubí y el hospital del Dispensari Veterinari del Vallès (Rubí 1).

**Cobertura de campos:**

| Campo | Cobertura |
|---|---|
| dirección | 100 % |
| coordenadas | 84,8 % |
| teléfono | 74,7 % |
| web | 62,0 % |
| email | 2,5 % |

## 4. Principales dificultades

1. **No se pudieron abrir páginas web.** La herramienta de descarga falló por DNS y el proxy del entorno bloquea Google Maps, OpenStreetMap/Overpass, Páginas Amarillas y rubi.cat. La verificación se hizo con un buscador que devuelve resúmenes de cada página, no la página original.
2. **No se consultó Google Maps.** Como alternativa se usó Overture Maps, que agrega fichas de Meta, Microsoft, Foursquare y localizadores oficiales de cadenas, y se descargó directamente del bucket público. De ahí salen las coordenadas y el límite municipal.
3. **La venta de alimentación casi nunca se publica por sucursal.** Los supermercados quedan como `no_verificado`, como exige el criterio.
4. **Los directorios están desfasados.** Hay fichas de más de 10 años, teléfonos contradictorios y negocios que han cerrado o cambiado de nombre (por ejemplo, El Regne Animal pasó a ser Rubí Nord, y VetClínic se concentró en un solo centro).
5. **La frontera con Sant Cugat no es evidente.** Tiendanimal, Leroy Merlin y un Caprabo aparecen como "Rubí" en directorios, pero están en Sant Cugat.
6. **Overture tiene errores.** Peluquerías clasificadas como veterinarios, direcciones mal asignadas a URLs oficiales y duplicados.

## 5. Información que no se ha podido contrastar

- **Disponibilidad de pienso de perro y gato** en 73 de los 76 registros no cerrados, incluidos todos los supermercados.
- **Actividad actual** de 25 registros en `pendiente_revision`, por ejemplo:
  - Boutique Animal: aviso de cierre temporal.
  - El Flamenco Rosa, Rubí Fauna, Cuida tus mascotas y Animal Paradis: solo aparecen en directorios antiguos.
  - Supermercados de barrio que solo constan en una fuente.
- **Coordenadas:** faltan en 12 registros. Uno es ALDI, que tiene dos puntos contradictorios a 750 m.
- **Teléfonos contradictorios:** Gosexpress, Fauna L'Amistat, Tot Pels Teus Amics y Condis Express Ctra. de Sabadell.
- **Charter:** Consum declara 7 franquicias en Rubí y solo se han identificado 6.
- **Coordenadas en general:** son los puntos de Overture del propio establecimiento, sin revisión visual sobre ortofoto.
- **Comercio:** no ha habido contacto telefónico ni visita. Ningún dato está validado por el propio comercio.

## 6. Nivel de confianza

- **Confirmado (alta confianza):** existencia y ubicación de las cadenas confirmadas en su web o localizador oficial (Lidl, ALDI, Dia, Consum/Charter, Condis Can Fatjó, Carrefour Express, Esclat, 2 de los 3 Mercadona), de los veterinarios con web propia y de las 3 tiendas `confirmado`.
- **Probable (confianza media):** el resto de registros `parcialmente_confirmado`. Existen y están en Rubí, pero falta el dato de alimentación o hay discrepancias menores.
- **Desconocido (confianza baja):** los 20 registros `no_verificado`. Deben tratarse como candidatos y no mostrarse en la app como establecimientos verificados.

**Valoración global:** la base sirve para el **piloto interno y para estimar la viabilidad**. El inventario de cadenas y veterinarios es fiable; las tiendas independientes están razonablemente cubiertas.

Antes de mostrar en la app el filtro "vende comida para perros o gatos", hace falta una **ronda de verificación telefónica o presencial**. Con los datos actuales solo podría mostrarse en 3 establecimientos.

## 7. Procedimiento recomendado para otros municipios

1. **Delimitar el municipio** con un polígono oficial (Overture/OSM o ICGC/IGN) y extraer todos los POI cartográficos que caen dentro. El script `scripts/01-03` es reutilizable.
2. **Completar las cadenas** con sus localizadores oficiales, una cadena cada vez.
3. **Completar los independientes** con el directorio de comercio local, el censo o las licencias municipales si existen, y búsquedas en castellano y catalán.
4. **Consolidar y deduplicar** por nombre, dirección, teléfono y URL, y aplicar una prueba de punto en polígono para la frontera municipal.
5. **Verificar por teléfono o en persona** con un guion fijo (¿pienso de perro?, ¿de gato?, ¿otros animales?, ¿horario?), empezando por las tiendas de mascotas, los veterinarios y una tienda de muestra por cadena. Registrar el resultado con `metodo_verificacion = comercio` o `manual`.
6. **Abrir la reclamación de fichas** (`reclamado_por_comercio`) como mecanismo de actualización continua.
7. **Conseguir acceso a Google Places API** (u otra fuente con estado de actividad) para reducir los `pendiente_revision`.
8. **Reutilizar `scripts/build_rubi_v1.py`.** Valida los dominios de valores y la coherencia entre campos, y mantiene los identificadores estables (`scripts/ids_rubi_v1.csv`): nunca se reutilizan.
