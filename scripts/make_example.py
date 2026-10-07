"""Create reproducible fictional data for testing, never overwrite differing raw input."""
from pathlib import Path
import json
import numpy as np
import pandas as pd

ROOT=Path(__file__).resolve().parents[1]


def example():
    rng=np.random.default_rng(42)
    n=1800
    category=rng.choice(['Accessories','Home','Office'],n,p=[.5,.3,.2])
    region=rng.choice(['North','South','West'],n,p=[.45,.35,.2])
    promo=rng.choice(['Yes','No'],n,p=[.4,.6])
    base=pd.Series(category).map({'Accessories':18.,'Home':45.,'Office':32.}).to_numpy()
    qty=rng.poisson(3+(category=='Accessories')*2+(promo=='Yes')*3)+1
    prices=np.maximum(1,base+rng.normal(0,4,n)).round(2)
    records=pd.DataFrame({'record_id':[f'R{i+1:05d}' for i in range(n)],
        'date':np.repeat(pd.date_range('2026-09-01',periods=30).strftime('%Y-%m-%d'),60),
        'product_category':category,'sales_region':region,'promoted':promo,
        'quantity':qty,'unit_price':prices,'currency':'INR',
        'item_id':[f'I{x:03d}' for x in rng.integers(1,101,n)],
        'product_url':[f'https://example.invalid/product/{x}' for x in rng.integers(1,101,n)],
        'notes':['fictional demonstration record']*n})
    records.loc[0:2,'quantity']=-1
    records.loc[20:23,'unit_price']=np.nan
    records.loc[30:31,'date']='invalid-date'
    records.loc[40:49,'product_category']=' accessories '
    records.loc[50:55,'sales_region']='NORTH '
    records.loc[60:64,'notes']=None
    records=pd.concat([records,records.iloc[:12]],ignore_index=True)
    cfg={'synthetic':True,'currency':'INR','row_key':'record_id','grain':'one fictional sales line',
         'date_formats':{'date':'%Y-%m-%d'},'nonnegative':['quantity','unit_price'],
         'roles':{'date':'date','quantity':'quantity','price':'unit_price','region':'sales_region','flag':'promoted'},
         'units':{'quantity':'units','unit_price':'INR per unit'},
         'assumptions':['Fictional seeded retail data, created solely to test the reporting format.',
                        'One row represents one fictional sales line; unit_price is per unit in INR.',
                        'Negative quantities are deliberately invalid test records, not returns.']}
    return records,cfg


if __name__=='__main__':
    frame,config=example()
    folder=ROOT/'data'/'raw';folder.mkdir(parents=True,exist_ok=True)
    outputs={folder/'example_retail.csv':frame.to_csv(index=False),
             folder/'example_retail.config.json':json.dumps(config,indent=2)+'\n'}
    for path,text in outputs.items():
        content=text.encode('utf-8')
        if path.exists() and path.read_bytes()!=content:
            raise SystemExit(f'Refusing to overwrite different existing source: {path}')
        if not path.exists():
            path.write_bytes(content)
    print(f'Created/verified {len(frame)} fictional raw rows. Seed=42. No real business data.')
