from flask import Flask, request, send_file, jsonify,render_template
from pptx import Presentation
from pptx.util import Inches, Pt, Cm
from pptx.dml.color import RGBColor
import os
import io,time
import base64
import matplotlib.pyplot as plt
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from PIL import Image
import zipfile
import os
import xml.etree.ElementTree as ET
import requests
import shutil

app = Flask(__name__)

def create_chart_xml(year1invest, year1return, year2invest, year2return,
                             year3invest, year3return, year4invest, year4return, year5invest, year5return):
    
    
    # barxml.txt dosyasından XML içeriğini oku
    with open('barxml.xml', 'r', encoding='utf-8') as file:
        xml_content = file.read()
    
    # Yıl değerlerini xml_content içinde değiştir
    xml_content = xml_content.replace('{year1invest}', str(year1invest))
    xml_content = xml_content.replace('{year1return}', str(year1return))
    xml_content = xml_content.replace('{year2invest}', str(year2invest))
    xml_content = xml_content.replace('{year2return}', str(year2return))
    xml_content = xml_content.replace('{year3invest}', str(year3invest))
    xml_content = xml_content.replace('{year3return}', str(year3return))
    xml_content = xml_content.replace('{year4invest}', str(year4invest))
    xml_content = xml_content.replace('{year4return}', str(year4return))
    xml_content = xml_content.replace('{year5invest}', str(year5invest))
    xml_content = xml_content.replace('{year5return}', str(year5return))
    
    return xml_content

def format_with_commas(value):
    try:
        # String değeri önce float veya int'e çevir
        num = float(value)
        
        # Virgüllü formatlama
        return f"{num:,.0f}"
    except ValueError:
        return "0"



def update_zip_with_new_xml(zip_path, output_zip_path, year1invest, year1return, year2invest, year2return,
                             year3invest, year3return, year4invest, year4return, year5invest, year5return):
    temp_dir = 'temp_zip'
    os.makedirs(temp_dir, exist_ok=True)

    # ZIP dosyasını çıkar
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    # Yeni XML içeriğini oluştur
    chart_xml_content = create_chart_xml(year1invest, year1return, year2invest, year2return,
                             year3invest, year3return, year4invest, year4return, year5invest, year5return)

    # Yeni içeriği chart1.xml olarak kaydet
    new_xml_path = os.path.join(temp_dir, 'ppt', 'charts', 'chart1.xml')
    with open(new_xml_path, 'w', encoding='utf-8') as xml_file:
        xml_file.write(chart_xml_content)

    # # <c:numCache> içerisine yeni verileri ekle
    # with open(new_xml_path, 'r+', encoding='utf-8') as xml_file:
    #     lines = xml_file.readlines()
    #     for index, line in enumerate(lines):
    #         if '<c:numCache>' in line:
    #             # Yeni verileri ekle
    #             lines.insert(index + 1, '    <c:ptCount val="10"/>\n')
    #             lines.insert(index + 2, f'    <c:pt idx="0"><c:v>{year1invest}</c:v></c:pt>\n')
    #             lines.insert(index + 3, f'    <c:pt idx="1"><c:v>{year1return}</c:v></c:pt>\n')
    #             lines.insert(index + 4, f'    <c:pt idx="2"><c:v>{year2invest}</c:v></c:pt>\n')
    #             lines.insert(index + 5, f'    <c:pt idx="3"><c:v>{year2return}</c:v></c:pt>\n')
    #             lines.insert(index + 6, f'    <c:pt idx="4"><c:v>{year3invest}</c:v></c:pt>\n')
    #             lines.insert(index + 7, f'    <c:pt idx="5"><c:v>{year3return}</c:v></c:pt>\n')
    #             lines.insert(index + 8, f'    <c:pt idx="6"><c:v>{year4invest}</c:v></c:pt>\n')
    #             lines.insert(index + 9, f'    <c:pt idx="7"><c:v>{year4return}</c:v></c:pt>\n')
    #             lines.insert(index + 10, f'    <c:pt idx="8"><c:v>{year5invest}</c:v></c:pt>\n')
    #             lines.insert(index + 11, f'    <c:pt idx="9"><c:v>{year5return}</c:v></c:pt>\n')
    #             break
    #     xml_file.seek(0)
    #     xml_file.writelines(lines)

    # Güncellenmiş dosyaları yeni ZIP dosyası olarak kaydet
    with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
        for foldername, subfolders, filenames in os.walk(temp_dir):
            for filename in filenames:
                filepath = os.path.join(foldername, filename)
                arcname = os.path.relpath(filepath, temp_dir)
                zip_ref.write(filepath, arcname)

    # Geçici dizini temizle
    shutil.rmtree(temp_dir)



def modify_slide_xml_and_image(zip_path, output_pptx_path,client_name,
                               itfinance='0',rpo='0',poa='0',cip='0',mspi='0',valmsl='0',valfqmr='0',valdcap='0',
                               valcifw='0',valoem='0',valbnft='0',valnpvv='0',valacd='0',valroi='0',valinvestment='0',valmonths='0',valhours='0',
                               year1invest='0',
                                year1return='0',year2invest='0',year2return='0',year3invest='0',
                                year3return='0',year4invest='0',year4return='0',year5invest='0',
                                year5return='0',costofdoingnothing1='0',itfinanceper='0',rpoper='0',poaper='0',cipper='0'
                                ,mspiper='0',valmslper='0',valfqmrper='0',valdcapper='0',
                               valcifwper='0',valoemper='0'):
    # Geçici çalışma dizinini oluştur
    temp_dir = 'temp_pptx'
    os.makedirs(temp_dir, exist_ok=True)

    # .pptx dosyasını aç ve dosyaları çıkar
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    # slides klasöründeki tüm slide XML dosyalarını bul ve işle
    slides_dir = os.path.join(temp_dir, 'ppt', 'slides')
    slide_files = [f for f in os.listdir(slides_dir) if f.startswith('slide') and f.endswith('.xml')]
    

    # Her bir slide dosyasını işle
    for slide_file in slide_files:
        slide_xml_path = os.path.join(slides_dir, slide_file)
        
        # XML dosyasını parse et
        tree = ET.parse(slide_xml_path)
        root = tree.getroot()

        # XML namespace tanımı (değiştirebilir)
        namespace = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}

        donutpercentvals = ''''''
        if rpoper!='0' and rpoper:
            donutpercentvals+='Raising Purchase Orders: We anticipate a {}% efficiency'.format(rpoper)
        # XML içeriğinde £XX,000 ifadelerini sırayla değiştir
        for elem in root.findall('.//a:t', namespace):
            if 'valclient' in elem.text:
                elem.text = elem.text.replace('valclient', client_name)
            if 'itfinance' in elem.text:
                elem.text = elem.text.replace('itfinance', format_with_commas(itfinance.replace('£','')))
            if 'valrpo' in elem.text:
                elem.text = elem.text.replace('valrpo', format_with_commas(rpo.replace('£','')))
            if 'valpoa' in elem.text:
                elem.text = elem.text.replace('valpoa', format_with_commas(poa.replace('£','')))
            if 'valcip' in elem.text:
                elem.text = elem.text.replace('valcip', format_with_commas(cip.replace('£','')))
            if 'mspi' in elem.text:
                elem.text = elem.text.replace('mspi', format_with_commas(mspi.replace('£','')))
            if 'valmsl' in elem.text:
                elem.text = elem.text.replace('valmsl', format_with_commas(valmsl.replace('£','')))
            if 'valfqmr' in elem.text:
                elem.text = elem.text.replace('valfqmr', format_with_commas(valfqmr.replace('£','')))
            if 'valdcap' in elem.text:
                elem.text = elem.text.replace('valdcap', format_with_commas(valdcap.replace('£','')))
            if 'valcifw' in elem.text:
                elem.text = elem.text.replace('valcifw', format_with_commas(valcifw.replace('£','')))
            if 'valoem' in elem.text:
                elem.text = elem.text.replace('valoem', format_with_commas(valoem.replace('£','')))
            if 'valbnft' in elem.text:
                elem.text = elem.text.replace('valbnft', format_with_commas(valbnft.replace('£','')))
            if 'valnpvv' in elem.text:
                elem.text = elem.text.replace('valnpvv', format_with_commas(valnpvv.replace('£','')))
            if 'valacd' in elem.text:
                elem.text = elem.text.replace('valacd', format_with_commas(valacd.replace('£','')))
            if 'valroi' in elem.text:
                elem.text = elem.text.replace('valroi', format_with_commas(valroi.replace('£','')))
            if 'valinvestment' in elem.text:
                elem.text = elem.text.replace('valinvestment', format_with_commas(valinvestment.replace('£','')))
            if 'valmonths' in elem.text:
                elem.text = elem.text.replace('valmonths', valmonths)
            if 'valhours' in elem.text:
                elem.text = elem.text.replace('valhours', valhours)
            if 'valcostof' in elem.text:
                elem.text = elem.text.replace('valcostof', format_with_commas(costofdoingnothing1.replace('£','')))
            if 'valdonutpercentvalues' in elem.text:
                elem.text = elem.text.replace('valdonutpercentvalues',donutpercentvals)
        zip_path = 'template.zip'  # Güncellemek istediğin template.zip
        output_zip_path = 'template.zip'  # Çıkış dosyasının adı

        # Yeni XML dosyasını oluştur ve ZIP dosyasını güncelle
        update_zip_with_new_xml(zip_path, output_zip_path,year1invest=year1invest,
                                year1return=year1return,year2invest=year2invest,year2return=year2return,year3invest=year3invest,
                                year3return=year3return,year4invest=year4invest,year4return=year4return,year5invest=year5invest,
                                year5return=year5return)

        # Güncellenmiş slide XML dosyasını kaydet
        tree.write(slide_xml_path, xml_declaration=True, encoding='UTF-8')


    

    # Güncellenmiş dosyaları tekrar ZIP yap
    with zipfile.ZipFile(output_pptx_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
        for foldername, subfolders, filenames in os.walk(temp_dir):
            for filename in filenames:
                filepath = os.path.join(foldername, filename)
                arcname = os.path.relpath(filepath, temp_dir)
                zip_ref.write(filepath, arcname)

    # Geçici çalışma dizinini temizle
    shutil.rmtree(temp_dir)

@app.route('/create-ppt', methods=['POST'])
def create_ppt():
    # 'client_name' parametresini POST isteği ile al
    data = request.get_json()
    client_name = data.get('client_name') or ""
    itfinance = data.get('itfinance') or ""
    rpo = data.get('rpo') or ""
    poa = data.get('poa') or ""
    cip = data.get('cip') or ""
    mspi = data.get('mspi') or ""
    valmsl = data.get('valmsl') or ""
    valfqmr = data.get('valfqmr') or ""
    valdcap = data.get('valdcap') or ""
    valcifw = data.get('valcifw') or ""
    valoem = data.get('valoem') or ""
    valbnft = data.get('valbnft') or ""
    valnpvv = data.get('valnpvv') or ""
    valacd = data.get('valacd') or ""
    valroi = data.get('valroi') or ""
    valinvestment = data.get('valinvestment') or ""
    valmonths = data.get('valmonths') or ""
    valhours = data.get('valhours') or ""
    year1total = data.get('year1total') or ""
    year1invest = data.get('year1invest') or ""
	
    year2otal = data.get('year2total') or ""
    year2invest = data.get('year2invest') or ""

    year3total = data.get('year3total') or ""
    year3invest = data.get('year3invest') or ""

    year4total = data.get('year4total') or ""
    year4invest = data.get('year4invest') or ""

    year5total = data.get('year5total') or ""
    year5invest = data.get('year5invest') or ""
    costofdoingnothing1 = data.get('costofdoingnothing1') or ""

    itfinanceper = data.get('itfinanceper') or ""
    rpoper = data.get('rpoper') or ""
    poaper = data.get('poaper') or ""
    cipper = data.get('cipper') or ""
    mspiper = data.get('mspiper') or ""
    valmslper = data.get('valmslper') or ""
    valfqmrper = data.get('valfqmrper') or ""
    valdcapper = data.get('valdcapper') or ""
    valcifwper = data.get('valcifwper') or ""
    valoemper = data.get('valoemper') or ""
    
    
    if not client_name:
        return "Error: 'client_name' parameter is required", 400





    zip_path = r"template.zip"  # Tam dosya yolunu girin
    output_pptx_path = r"output.pptx"  # Çıkış dosyasının yolunu belirtin

    modify_slide_xml_and_image(zip_path, output_pptx_path,client_name,itfinance,rpo,poa,cip,mspi,valmsl,valfqmr,valdcap,
                               valcifw,valoem,valbnft,valnpvv,valacd,valroi,valinvestment,valmonths,valhours,
                               year1invest=year1invest,
                                year1return=year1total,year2invest=year2invest,year2return=year2otal,year3invest=year3invest,
                                year3return=year3total,year4invest=year4invest,year4return=year4total,year5invest=year5invest,
                                year5return=year5total,costofdoingnothing1=costofdoingnothing1)
    
    zip_path = 'template.zip'  # Güncellemek istediğin template.zip
    output_zip_path = 'template.zip'  # Çıkış dosyasının adı



    pptx_io = io.BytesIO()
    with open(output_pptx_path, 'rb') as f:
        pptx_io.write(f.read())
    pptx_io.seek(0)

    # Dosyayı base64 formatında encode et
    pptx_base64 = base64.b64encode(pptx_io.read()).decode('utf-8')

    # JSON formatında base64 ile encode edilmiş dosyayı döndür
    return jsonify({
        'file_name': 'presentation.pptx',
        'file_content': pptx_base64
    })




if __name__ == "__main__":
    app.run(host='0.0.0.0', port=8080)
