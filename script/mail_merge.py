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
            else:
                data[i] = str(data[i])
                
        f = open('template.docx', 'rb')
        doc = Document(f)
        ## General Information
        doc.tables[0].cell(0, 1).text = data['doctor_name']
        doc.tables[0].cell(1, 1).text = data['doctor_department']
        doc.tables[0].cell(2, 1).text = str(data['doctor_sip'])
        doc.tables[0].cell(3, 1).text = str(data['doctor_str'])
        doc.tables[0].cell(0, 3).text = data['name']
        doc.tables[0].cell(1, 3).text = data['gender']
        doc.tables[0].cell(2, 3).text = str(data['dob'])
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
        doc.paragraphs[4].text = str(data['diagnosis']) if str(data['diagnosis']) != '' else ''
        doc.paragraphs[4].runs[0].font.size= Pt(8)
        
        doc.paragraphs[6].text = data['icdx']
        doc.paragraphs[6].runs[0].font.size= Pt(8)

        doc.paragraphs[8].text = str(data['suggestion']) if str(data['suggestion']) != '' else '' 
        doc.paragraphs[8].runs[0].font.size= Pt(8)
        
        ## Consultation Detail
        doc.tables[1].cell(1, 1).text = str(data['consult_id']).replace('.0', '')
        doc.tables[1].cell(1, 2).text = str(data['claim_id']).replace('.0', '')
        doc.tables[1].cell(2, 2).text = str(data['claim_id_rx']).replace('.0', '')
        doc.tables[1].cell(1, 3).text = pd.to_datetime(data['date']).strftime('%d/%m/%Y') + ' ' + data['time']
        doc.tables[1].cell(1, 4).text = 'IDR '+ str(data['consult_fee']).replace('.0', '')
        doc.tables[1].cell(1, 7).text = 'IDR '+ str(data['consult_fee']).replace('.0', '')
        doc.tables[1].cell(2, 1).text = str(data['order_id']).replace('.0', '')
        doc.tables[1].cell(2, 3).text = '' if str(data['order_created_date']).lower() in ('nat', 'nan', '') else pd.to_datetime(data['order_created_date'], errors='coerce').strftime('%d/%m/%Y') + ' ' + str(data['order_created_time'])
        
        
        doc.tables[1].cell(1, 1).paragraphs[0].runs[0].font.size= Pt(8)
        doc.tables[1].cell(1, 2).paragraphs[0].runs[0].font.size= Pt(8)
        doc.tables[1].cell(2, 2).paragraphs[0].runs[0].font.size= Pt(8)
        doc.tables[1].cell(1, 3).paragraphs[0].runs[0].font.size= Pt(8)
        doc.tables[1].cell(1, 4).paragraphs[0].runs[0].font.size= Pt(8)
        doc.tables[1].cell(1, 7).paragraphs[0].runs[0].font.size= Pt(8)
        doc.tables[1].cell(2, 1).paragraphs[0].runs[0].font.size= Pt(8)
        doc.tables[1].cell(2, 3).paragraphs[0].runs[0].font.size= Pt(8)
        
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
        
        doc.tables[1].cell(16, 4).text = 'IDR 0' if data['deliv_coverage_by_Insurance'] =='' else 'IDR '+ str(data['deliv_coverage_by_Insurance']).replace('.0', '')
        doc.tables[1].cell(16, 7).text = 'IDR 0' if data['deliv_coverage_by_Insurance'] =='' else 'IDR '+ str(data['deliv_coverage_by_Insurance']).replace('.0', '')
        doc.tables[1].cell(18, 7).text = 'IDR '+ str(data['total']).replace('.0', '')
        doc.tables[1].cell(19, 7).text = 'IDR '+ str(data['total_consult_+_rx']).replace('.0', '')
        doc.tables[1].cell(20, 7).text = 'IDR '+ str(float(data['total']) - float(data['total_consult_+_rx'])).replace('.0', '')
        
        doc.tables[1].cell(16, 4).paragraphs[0].runs[0].font.size= Pt(8)
        doc.tables[1].cell(16, 7).paragraphs[0].runs[0].font.size= Pt(8)
        doc.tables[1].cell(18, 7).paragraphs[0].runs[0].font.size= Pt(8)
        doc.tables[1].cell(19, 7).paragraphs[0].runs[0].font.size= Pt(8)
        doc.tables[1].cell(20, 7).paragraphs[0].runs[0].font.size= Pt(8)
        
        doc.tables[1].cell(18, 7).paragraphs[0].runs[0].bold = True
        doc.tables[1].cell(19, 7).paragraphs[0].runs[0].bold = True
        doc.tables[1].cell(20, 7).paragraphs[0].runs[0].bold = True
        
        name = data['name'].replace('/', ' ').title()
        consult_on = pd.to_datetime(data['date'], errors = 'coerce').strftime('%Y%m%d') #+'_'+str(data['time']).replace(':', '')
        consult_id = str(data['consult_id']).replace('.0', '')
        agg = data['aggregator_name']
        payor = data['payor']

        if not os.path.exists(f'.tmp/{self.label}/{agg}/{payor}'):
            os.makedirs(f'.tmp/{self.label}/{agg}/{payor}')
        
        tmp_name = f'.tmp/{self.label}/{agg}/{payor}/{consult_on}_{name}_Consultation_Receipt_{consult_id}.docx'
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
        fls = sorted(fls)
        unique_agg = np.unique([i.split('/')[2] for i in fls])
        unique_pay = np.unique([i.split('/')[3] for i in fls])
        
        success = []
        
        for i in unique_agg:
            for j in unique_pay:
                ij = [x for x in fls if x.split('/')[2] == i and x.split('/')[3] == j]
                k = 1
                for files in chunks(ij, self.file_per_zip):
                    files = sorted(files)
                    with ZipFile(f'output/{self.label}/{i}_{j}_{k}.zip','w') as zip: 
                        for file_each in files:
                            zip.write(file_each, file_each.split('/')[-1]) 
                            success.append([os.path.basename(file_each),f'{i}_{j}_{k}.zip'])
                    k = k + 1


        self.errors.to_excel(f'output/{self.label}/errors.xlsx', index = False)
        pd.DataFrame(success).to_excel(f'output/{self.label}/success.xlsx', index = False)
        
        #self.dataset[['date', 'name', 'card_no', 'gdt_id', 
        #            'consult_id','icdx','gpsp','claim_id', 'consult_price','prescription_fee', 'total']]

        fls2 = glob.glob(f'output/{self.label}/*')
        
        with ZipFile(f'output/{self.label}.zip','w') as zip:
            for file_each in fls2:
                zip.write(file_each, file_each.split('/')[-1]) 
        shutil.rmtree(f'output/{self.label}')

    def run(self):
        if os.path.exists(f'.tmp/{self.label}'):
            shutil.rmtree(f'.tmp/{self.label}')
        
        errors = []
        for index, row in tqdm(self.dataset.iterrows()):
            try:
                self.row_to_pdf(row)
            except Exception as e:
                errors.append(row)
                print(e)

        self.errors = pd.DataFrame(errors)

        self.chunk_and_zip()


def run_mail_merge(file, label, sheet_name = 'data', file_per_zip = 100):
    df = pd.read_excel(file)
    MailMerge(df, label, file_per_zip = file_per_zip).run()
    