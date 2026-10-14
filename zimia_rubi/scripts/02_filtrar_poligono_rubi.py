# Paso 2: filtra los lugares de places.parquet que caen dentro del polígono municipal de Rubí (div.parquet)
# y guarda inside.pkl. Imprime el recuento por categoría.
import pyarrow.parquet as pq, shapely, collections, json
div=[r for r in pq.read_table('div.parquet').to_pylist() if (r['names'] or {}).get('primary')=='Rubí'][0]
poly=shapely.from_wkb(div['geometry']); print(poly.geom_type, poly.area)
t=pq.read_table('places.parquet'); print(t.schema)
rows=t.to_pylist(); inside=[]
for r in rows:
    p=shapely.from_wkb(r['geometry'])
    if poly.contains(p): r['_lon'],r['_lat']=p.x,p.y; inside.append(r)
print(len(inside))
c=collections.Counter((r['basic_category'],(r['taxonomy'] or {}).get('primary')) for r in inside)
for k,v in c.most_common(400): print(v,k)
import pickle; pickle.dump(inside,open('inside.pkl','wb'))
