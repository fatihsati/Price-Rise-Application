from docx import Document
from docx.shared import Pt
import math



def price_text_format(price, currency_unit):

    price = str(price) + f'.00{currency_unit}'
    return price
    

def zam_to_cell(price, zam_orani, round_up, currency_unit):
    new_price = int(price + (price * zam_orani / 100))
    # en yakin 5'in katina yuvarla
    if new_price % 5 != 0:
        if round_up == True:
            new_price = (math.floor(new_price/5)*5)+5
        else:
            new_price = math.floor(new_price/5)*5

    new_price = price_text_format(new_price, currency_unit)
    return new_price



def change_doc(docname, zam_orani, round_up, currency_unit):
    # docname, zam_orani, round_up = get_inputs()
    doc = Document(docname)
    table = doc.tables[0]
    for row in table.rows[1:]:    
        for cell in row.cells[1:]:
            try:
                cell_value = cell.text.split('.')[0] 
                if currency_unit in cell_value:
                    price = cell_value[1:].strip()
                else:
                    price = cell_value
                cell.text = zam_to_cell(int(price), zam_orani, round_up, currency_unit)
                                       
            except:
                continue
    return doc


