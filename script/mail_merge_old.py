from script.utils import chunks
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import subprocess
from docx.shared import Pt
import pandas as pd
import numpy as np
import os
import shutil
from tqdm import tqdm
import glob
from zipfile import ZipFile 
import concurrent

class MailMerge():
    def __init__(self, dataset ,label, file_per_zip = 100, parallel = False):
        self.dataset = dataset
        self.label = label
        self.file_per_zip = file_per_zip
        self.multicore = parallel
        
    def row_to_pdf(self, data):
        for i in data.keys():
            if str(data.get(i)) == 'nan':
                data[i] = ''
                
        f = open('template.docx', 'rb')
        doc = Document(f)
        
        ## General Information
        doc.tables[0].cell(0, 1).text = data['doctor_name']
        doc.tables[0].cell(1, 1).text = data['doctor_department']
        doc.tables[0].cell(2, 1).text = data['doctor_sip']
        doc.tables[0].cell(3, 1).text = data['doctor_str']
        doc.tables[0].cell(0, 3).text = data['name']
        doc.tables[0].cell(1, 3).text = data['gender']
        doc.tables[0].cell(2, 3).text = data['dob']
        doc.tables[0].cell(3, 3).text = str(data['card_no']).replace('.0', '')
        doc.tables[0].cell(4, 3).text = data['payor']
        doc.tables[0].cell(5, 3).text = data['corporate']
        
        for i in range(6):
            if i < 4:
                doc.tables[0].cell(i, 1).paragraphs[0].runs[0].font.size = Pt(8)
            doc.tables[0].cell(i, 3).paragraphs[0].runs[0].font.size = Pt(8)
        
        ## complaint & diagnosis
        doc.paragraphs[2].text = data['chief_complaints']
        doc.paragraphs[2].runs[0].font.size= Pt(8)
        doc.paragraphs[4].text = str(data['diagnosis']) if str(data['diagnosis']) != '' else '' + (', ' + str(data['suggestion'])) if str(data['suggestion']) != '' else '' 
        doc.paragraphs[4].runs[0].font.size= Pt(8)
        
        doc.paragraphs[6].text = data['icdx']
        doc.paragraphs[6].runs[0].font.size= Pt(8)
        
        ## Consultation Detail
        doc.tables[1].cell(1, 1).text = str(data['consult_id']).replace('.0', '')
        doc.tables[1].cell(1, 2).text = str(data['claim_id']).replace('.0', '')
        doc.tables[1].cell(1, 3).text = data['date'].strftime('%d/%m/%Y') + ' ' + data['time']
        doc.tables[1].cell(1, 4).text = 'IDR '+ str(data['consult_fee']).replace('.0', '')
        doc.tables[1].cell(1, 7).text = 'IDR '+ str(data['consult_fee']).replace('.0', '')
        doc.tables[1].cell(2, 1).text = str(data['order_id']).replace('.0', '')
        
        
        doc.tables[1].cell(1, 1).paragraphs[0].runs[0].font.size= Pt(8)
        doc.tables[1].cell(1, 2).paragraphs[0].runs[0].font.size= Pt(8)
        doc.tables[1].cell(1, 3).paragraphs[0].runs[0].font.size= Pt(8)
        doc.tables[1].cell(1, 4).paragraphs[0].runs[0].font.size= Pt(8)
        doc.tables[1].cell(1, 7).paragraphs[0].runs[0].font.size= Pt(8)
        doc.tables[1].cell(2, 1).paragraphs[0].runs[0].font.size= Pt(8)
        
        for i in range(1,13):
            j = i + 2
            if str(data[f'obat_{i}']) != '':
                doc.tables[1].cell(j, 0).text = data[f'obat_{i}']
                doc.tables[1].cell(j, 0).paragraphs[0].alignment = 2
                doc.tables[1].cell(j, 4).text = 'IDR '+ str(data[f'total_{i}']).replace('.0', '')
                doc.tables[1].cell(j, 5).text = str(data[f'jumlah_{i}']).replace('.0', '')
                doc.tables[1].cell(j, 6).text = data[f'unit_obat_{i}']
                doc.tables[1].cell(j, 7).text = 'IDR '+ str(data[f'total_{i}']).replace('.0', '')
        
                doc.tables[1].cell(j, 6).paragraphs[0].runs[0].italic = True
                
                doc.tables[1].cell(j, 0).paragraphs[0].runs[0].font.size= Pt(8)
                doc.tables[1].cell(j, 4).paragraphs[0].runs[0].font.size= Pt(8)
                doc.tables[1].cell(j, 5).paragraphs[0].runs[0].font.size= Pt(8)
                doc.tables[1].cell(j, 6).paragraphs[0].runs[0].font.size= Pt(8)
                doc.tables[1].cell(j, 7).paragraphs[0].runs[0].font.size= Pt(8)
        
        doc.tables[1].cell(16, 4).text = 'IDR 0' if data['excess_delivery_fee'] =='' else 'IDR '+ str(data['excess_delivery_fee']).replace('.0', '')
        doc.tables[1].cell(16, 7).text = 'IDR 0' if data['excess_delivery_fee'] =='' else 'IDR '+ str(data['excess_delivery_fee']).replace('.0', '')
        doc.tables[1].cell(18, 7).text = 'IDR '+ str(data['total']).replace('.0', '')
        excess =  0 if  data['rx_excess'] == '' else float(data['rx_excess']) + 0 if  data['excess_consult'] == '' else data['excess_consult']
        doc.tables[1].cell(19, 7).text = 'IDR '+ str(data['total'] - excess).replace('.0', '')
        doc.tables[1].cell(20, 7).text = 'IDR '+ str(excess).replace('.0', '')
        
        doc.tables[1].cell(16, 4).paragraphs[0].runs[0].font.size= Pt(8)
        doc.tables[1].cell(16, 7).paragraphs[0].runs[0].font.size= Pt(8)
        doc.tables[1].cell(18, 7).paragraphs[0].runs[0].font.size= Pt(8)
        doc.tables[1].cell(19, 7).paragraphs[0].runs[0].font.size= Pt(8)
        doc.tables[1].cell(20, 7).paragraphs[0].runs[0].font.size= Pt(8)
        
        doc.tables[1].cell(18, 7).paragraphs[0].runs[0].bold = True
        doc.tables[1].cell(19, 7).paragraphs[0].runs[0].bold = True
        doc.tables[1].cell(20, 7).paragraphs[0].runs[0].bold = True
        
        name = data['name'].replace('/', ' ')
        consult_on = data['date'].strftime('%Y%m%d') +'_'+str(data['time']).replace(':', '')
        consult_id = str(data['consult_id']).replace('.0', '')
        agg = data['aggregator_name']
        payor = data['payor']
        
        if not os.path.exists(f'.tmp/{self.label}/{agg}/{payor}'):
            os.makedirs(f'.tmp/{self.label}/{agg}/{payor}')
        
        
        tmp_name = f'.tmp/{self.label}/{agg}/{payor}/{consult_on}_Consultation_Receipt_{name}_{consult_id}.docx'
        doc.save(tmp_name)
        f.close()
        
        subprocess.run(['libreoffice', '--convert-to', 'pdf' ,
                        tmp_name, '--outdir', f'.tmp/{self.label}/{agg}/{payor}']
                       ,stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL
                    )
        os.remove(tmp_name)
        
        
    def chunk_and_zip(self):
        if not os.path.exists(f'output/{self.label}'):
            os.makedirs(f'output/{self.label}')
        else:
            shutil.rmtree(f'output/{self.label}')
            raise Exception("Error: Anda sedang menjalankan mail merge dengan label yang sama, silahkan ganti salah satu label")
    
        if os.path.exists(f'output/{self.label}.zip'):
            os.remove(f'output/{self.label}.zip')
            
        fls = glob.glob(f'.tmp/{self.label}/*/*/*')
        unique_agg = np.unique([i.split('/')[2] for i in fls])
        unique_pay = np.unique([i.split('/')[3] for i in fls])
        for i in unique_agg:
            for j in unique_pay:
                ij = [x for x in fls if x.split('/')[2] == i and x.split('/')[3] == j]
                k = 1
                for files in chunks(ij, self.file_per_zip):
                    with ZipFile(f'output/{self.label}/{i}_{j}_{k}.zip','w') as zip: 
                        for file_each in files:
                            zip.write(file_each, file_each.split('/')[-1]) 
                    k = k + 1
                    
        fls2 = glob.glob(f'output/{self.label}/*')
        with ZipFile(f'output/{self.label}.zip','w') as zip:
            for file_each in fls2:
                zip.write(file_each, file_each.split('/')[-1]) 
        shutil.rmtree(f'output/{self.label}')

    def run(self):
        if os.path.exists(f'.tmp/{self.label}'):
            shutil.rmtree(f'.tmp/{self.label}')
        
        if not self.multicore:
            for index, row in tqdm(self.dataset.iterrows()):
                self.row_to_pdf(row)
        else:
            executor = concurrent.futures.ProcessPoolExecutor(4)
            futures = [executor.submit(self.row_to_pdf, item) for item in self.dataset.to_dict(orient='records')]
            concurrent.futures.wait(futures)
        
        self.chunk_and_zip()


def run_mail_merge(file, label, sheet_name = 'data', file_per_zip = 100, parallel = False):
    df = pd.read_excel(file, sheet_name = sheet_name)
    MailMerge(df, label, file_per_zip = file_per_zip, parallel = parallel).run()
    