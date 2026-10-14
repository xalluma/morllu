# Paso 3: preselección de candidatos (categorías y palabras clave en castellano y catalán) desde inside.pkl.
# Salida: cand.pkl y listado JSON por stdout. La revisión y verificación posterior es manual (ver build_rubi_v1.py).
import pickle,re,json
inside=pickle.load(open('inside.pkl','rb'))
CATS={'veterinarian','pet_store','animal_or_pet_service','pet_boarding','pet_breeder','pet_groomer','grocery_store','supermarket','convenience_store','discount_store','shopping','nursery_and_gardening_store','farm','health_food_store','vitamin_and_supplement_store','food_and_beverage_store','hardware_store','home_improvement_store','department_store','shopping_mall','produce_store','hypermarket','agricultural_supply_store','korean_grocery_store','home_and_garden','garden_center'}
KW=re.compile(r'mascot|animal|pet|vet|pinso|pienso|gos|gat|perr|agro|agr[ií]col|garden|jard|herbol|natur|bio|eco|mercadona|lidl|aldi|carrefour|caprabo|bonpreu|bon preu|condis|consum|dia\b|bonàrea|bonarea|alcampo|eroski|spar|sorli|esclat|suma|plusfresc|ametller|veritas|mercat|hiper|super|leroy|bricomart|verdecora|kiwoko|tiendanimal|zooplus|miscota|guau|miau|cuc|can |ferret',re.I)
out=[]
for r in inside:
    tp=(r['taxonomy'] or {}).get('primary'); bc=r['basic_category']; n=(r['names'] or {}).get('primary') or ''
    alts=(r['taxonomy'] or {}).get('alternates') or []
    if tp in CATS or bc in ('animal_or_pet_service','animal_and_pet_store') or KW.search(n) or set(alts)&CATS:
        a=(r['addresses'] or [{}])[0] if r['addresses'] else {}
        b=((r['brand'] or {}).get('names') or {}).get('primary') if r['brand'] else None
        out.append(dict(id=r['id'],name=n,tp=tp,alts=alts,brand=b,addr=a.get('freeform'),pc=a.get('postcode'),loc=a.get('locality'),lat=round(r['_lat'],6),lon=round(r['_lon'],6),ph=r['phones'],web=r['websites'],em=r['emails'],soc=r['socials'],conf=round(r['confidence'] or 0,2),op=r['operating_status'],src=[(s['dataset'],s.get('update_time')) for s in r['sources']]))
pickle.dump(out,open('cand.pkl','wb'))
for o in sorted(out,key=lambda x:(x['tp'] or '',x['name'])): print(json.dumps(o,ensure_ascii=False))
print(len(out))
