from zipfile import ZipFile 
import numpy as np

def chunks(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n] 

def chunk_and_zip(outdir):
    fls = glob.glob(f'{outdir}/*/*/*')
    unique_agg = np.unique([i.split('/')[1] for i in fls])
    unique_pay = np.unique([i.split('/')[2] for i in fls])
    for i in unique_agg:
        for j in unique_pay:
            ij = [x for x in fls if x.split('/')[1] == i and x.split('/')[2] == j]
            k = 1
            for files in chunks(ij, 100):
                with ZipFile(f'output2/{i}_{j}_{k}.zip','w') as zip: 
                    for file_each in files:
                        zip.write(file_each) 
                k = k + 1