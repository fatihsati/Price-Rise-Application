from docx import Document
from docx.shared import Pt
import math



def format_fiyat(fiyat, para_birimi):

    fiyat = str(fiyat) + f'.00{para_birimi}'
    return fiyat
    

def zam_to_cell(fiyat, zam_orani, round_up, para_birimi):
    yeni_fiyat = int(fiyat + (fiyat * zam_orani / 100))
    # en yakin 5'in katina yuvarla
    if yeni_fiyat % 5 != 0:
        if round_up == True:
            yeni_fiyat = (math.floor(yeni_fiyat/5)*5)+5
        else:
            yeni_fiyat = math.floor(yeni_fiyat/5)*5

    yeni_fiyat = format_fiyat(yeni_fiyat, para_birimi)
    return yeni_fiyat



def change_doc(docname, zam_orani, round_up, para_birimi):
    # docname, zam_orani, round_up = get_inputs()
    doc = Document(docname)
    table = doc.tables[0]
    for row in table.rows[1:]:    
        for cell in row.cells[1:]:
            try:
                cell_value = cell.text.split('.')[0] 
                if para_birimi in cell_value:
                    fiyat = cell_value[1:].strip()
                else:
                    fiyat = cell_value
                cell.text = zam_to_cell(int(fiyat), zam_orani, round_up, para_birimi)
                
            except:
                continue
    return doc


