# Paso 1: descarga de Overture (places y divisions) para un bbox amplio alrededor de Rubí.
# Uso: python3 01_extraer_overture.py places place places.parquet
#      python3 01_extraer_overture.py divisions division_area div.parquet
import pyarrow.fs as pfs, pyarrow.dataset as ds, pyarrow.parquet as pq, os, sys
fs=pfs.S3FileSystem(anonymous=True, region='us-west-2', proxy_options=os.environ['HTTPS_PROXY'])
R="overturemaps-us-west-2/release/2026-09-23.0"
def get(theme,typ,out):
    d=ds.dataset(f"{R}/theme={theme}/type={typ}",filesystem=fs,format="parquet")
    f=(ds.field("bbox","xmin")<2.12)&(ds.field("bbox","xmax")>1.95)&(ds.field("bbox","ymin")<41.56)&(ds.field("bbox","ymax")>41.42)
    t=d.to_table(filter=f)
    pq.write_table(t,out); print(typ,t.num_rows)
get(sys.argv[1],sys.argv[2],sys.argv[3])
