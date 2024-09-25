from flask import Flask, request, send_file, jsonify,render_template
from pptx import Presentation
from pptx.util import Inches, Pt, Cm
from pptx.dml.color import RGBColor
from bs4 import BeautifulSoup
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
from pptx.enum.text import PP_ALIGN

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

def create_donut_xml(donutit,donutrpo,donutpoa,
                             donutdcap,donutcip,donutmspi,donutmsl,donutfqmr,donutcifw,donutoem):
    
    
    # barxml.txt dosyasından XML içeriğini oku
    with open('donutxml.xml', 'r', encoding='utf-8') as file:
        xml_content = file.read()
    
    # Yıl değerlerini xml_content içinde değiştir
    xml_content = xml_content.replace('{donutrpo}', str(donutrpo))
    xml_content = xml_content.replace('{donutpoa}', str(donutpoa))
    xml_content = xml_content.replace('{donutcip}', str(donutcip))
    xml_content = xml_content.replace('{donutmspi}', str(donutmspi))
    xml_content = xml_content.replace('{donutmsl}', str(donutmsl))
    xml_content = xml_content.replace('{donutfqmr}', str(donutfqmr))
    xml_content = xml_content.replace('{donutdcap}', str(donutdcap))
    xml_content = xml_content.replace('{donutcifw}', str(donutcifw))
    xml_content = xml_content.replace('{donutit}', str(donutit))
    xml_content = xml_content.replace('{donutoem}', str(donutoem))
    
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
                             year3invest, year3return, year4invest, year4return, year5invest, year5return,donutit,donutrpo,donutpoa,
                             donutdcap,donutcip,donutmspi,donutmsl,donutfqmr,donutcifw,donutoem):
    temp_dir = 'temp_zip'
    os.makedirs(temp_dir, exist_ok=True)

    # ZIP dosyasını çıkar
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    # Yeni XML içeriğini oluştur
    chart_xml_content = create_chart_xml(year1invest, year1return, year2invest, year2return,
                             year3invest, year3return, year4invest, year4return, year5invest, year5return)
    
    donut_xml_content = create_donut_xml(donutit,donutrpo,donutpoa,
                             donutdcap,donutcip,donutmspi,donutmsl,donutfqmr,donutcifw,donutoem)

    new_xml_path = os.path.join(temp_dir, 'ppt', 'charts', 'chart2.xml')
    with open(new_xml_path, 'w', encoding='utf-8') as xml_file:
        xml_file.write(chart_xml_content)

    new_xml_path = os.path.join(temp_dir, 'ppt', 'charts', 'chart1.xml')
    with open(new_xml_path, 'w', encoding='utf-8') as xml_file:
        xml_file.write(donut_xml_content)

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
                               valcifwper='0',valoemper='0',
                               donutit='0',donutrpo='0',donutpoa='0',donutdcap='0',donutcip='0',donutmspi='0',donutmsl='0',
                               donutfqmr='0',donutcifw='0',donutoem='0'):
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
            donutpercentvals+='\nRaising Purchase Orders: We anticipate a {}% efficiency'.format(rpoper)
        if itfinanceper!='0' and itfinanceper:
            donutpercentvals+='\nIT finance systems: We anticipate a {}% efficiency'.format(itfinanceper)
        if poaper!='0' and poaper:
            donutpercentvals+='\nPurchase Order Approvals: We anticipate a {}% efficiency'.format(poaper)
        if cipper!='0' and cipper:
            donutpercentvals+='\nCoding invoice processes: We anticipate a {}% efficiency'.format(cipper)
        if mspiper!='0' and mspiper:
            donutpercentvals+='\nManagement of Supplier and Purchase Invoices: We anticipate a {}% efficiency'.format(mspiper)
        if valmslper!='0' and valmslper:
            donutpercentvals+='\nManaging Maverick Spend & Spend Leakage: We anticipate a {}% efficiency'.format(valmslper)
        if valfqmrper!='0' and valfqmrper:
            donutpercentvals+='\nFinance Query Management and Dashboard Reporting: We anticipate a {}% efficiency'.format(valfqmrper)
        if valdcapper!='0' and valdcapper:
            donutpercentvals+='\nDebt Collection Administration Processes: We anticipate a {}% efficiency'.format(valdcapper)
        if valcifwper!='0' and valcifwper:
            donutpercentvals+='\nCustomer Invoicing & Finance Workflow Management: We anticipate a {}% efficiency'.format(valcifwper)
        if valoemper!='0' and valoemper:
            donutpercentvals+='\nOnline expense management: We anticipate a {}% efficiency'.format(valoemper)
        
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
                valhours+='h'
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

    donutit = data.get('donutit') or "0"
    donutrpo = data.get('donutrpo') or "0"
    donutpoa = data.get('donutpoa') or "0"
    donutdcap = data.get('donutdcap') or "0"
    donutcip = data.get('donutcip') or "0"
    donutmspi = data.get('donutmspi') or "0"
    donutmsl = data.get('donutmsl') or "0"
    donutfqmr = data.get('donutfqmr') or "0"
    donutcifw = data.get('donutcifw') or "0"
    donutoem = data.get('donutoem') or "0"
    
    if not client_name:
        return "Error: 'client_name' parameter is required", 400

    zip_dosya = 'template.zip'
    gecici_zip_dosya = 'temp_template.zip'

    



    zip_path = r"template.zip"  # Tam dosya yolunu girin
    output_pptx_path = r"output.pptx"  # Çıkış dosyasının yolunu belirtin

    modify_slide_xml_and_image(zip_path, output_pptx_path,client_name,itfinance,rpo,poa,cip,mspi,valmsl,valfqmr,valdcap,
                               valcifw,valoem,valbnft,valnpvv,valacd,valroi,valinvestment,valmonths,valhours,
                               year1invest=year1invest,
                                year1return=year1total,year2invest=year2invest,year2return=year2otal,year3invest=year3invest,
                                year3return=year3total,year4invest=year4invest,year4return=year4total,year5invest=year5invest,
                                year5return=year5total,costofdoingnothing1=costofdoingnothing1,itfinanceper=itfinanceper,rpoper=rpoper,
                                poaper=poaper,cipper=cipper,mspiper=mspiper,valmslper=valmslper,valfqmrper=valfqmrper,valdcapper=valdcapper,
                                valcifwper=valcifwper,valoemper=valoemper,donutit=donutit,donutrpo=donutrpo,donutpoa=donutpoa,
                                donutdcap=donutdcap,donutcip=donutcip,donutmspi=donutmspi,donutmsl=donutmsl,donutfqmr=donutfqmr,
                                donutcifw=donutcifw,donutoem=donutoem)
    
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
