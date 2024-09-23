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

def modify_slide_xml_and_image(zip_path, output_pptx_path,client_name,
                               itfinance=0,rpo=0,poa=0,cip=0,mspi=0,valmsl=0,valfqmr=0,valdcap=0,
                               valcifw=0,valoem=0,valbnft=0,valnpvv=0,valacd=0,valroi=0,valinvestment=0,valmonths=0):
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

        # XML içeriğinde £XX,000 ifadelerini sırayla değiştir
        for elem in root.findall('.//a:t', namespace):
            if 'valclient' in elem.text:
                elem.text = elem.text.replace('valclient', client_name)
            if 'itfinance' in elem.text:
                elem.text = elem.text.replace('itfinance', itfinance.replace("£", "").replace(" ", ""))
            if 'valrpo' in elem.text:
                elem.text = elem.text.replace('valrpo', rpo.replace("£", "").replace(" ", ""))
            if 'valpoa' in elem.text:
                elem.text = elem.text.replace('valpoa', poa.replace("£", "").replace(" ", ""))
            if 'valcip' in elem.text:
                elem.text = elem.text.replace('valcip', cip.replace("£", "").replace(" ", ""))
            if 'mspi' in elem.text:
                elem.text = elem.text.replace('mspi', mspi.replace("£", "").replace(" ", ""))
            if 'valmsl' in elem.text:
                elem.text = elem.text.replace('valmsl', valmsl.replace("£", "").replace(" ", ""))
            if 'valfqmr' in elem.text:
                elem.text = elem.text.replace('valfqmr', valfqmr.replace("£", "").replace(" ", ""))
            if 'valdcap' in elem.text:
                elem.text = elem.text.replace('valdcap', valdcap.replace("£", "").replace(" ", ""))
            if 'valcifw' in elem.text:
                elem.text = elem.text.replace('valcifw', valcifw.replace("£", "").replace(" ", ""))
            if 'valoem' in elem.text:
                elem.text = elem.text.replace('valoem', valoem.replace("£", "").replace(" ", ""))
            if 'valbnft' in elem.text:
                elem.text = elem.text.replace('valbnft', valbnft.replace("£", "").replace(" ", ""))
            if 'valnpvv' in elem.text:
                elem.text = elem.text.replace('valnpvv', valnpvv.replace("£", "").replace(" ", ""))
            if 'valacd' in elem.text:
                elem.text = elem.text.replace('valacd', valacd.replace("£", "").replace(" ", ""))
            if 'valroi' in elem.text:
                elem.text = elem.text.replace('valroi', valroi)
            if 'valinvestment' in elem.text:
                elem.text = elem.text.replace('valinvestment', valinvestment.replace("£", "").replace(" ", ""))
            if 'valmonths' in elem.text:
                elem.text = elem.text.replace('valmonths', valmonths)

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
    client_name = data.get('client_name')
    itfinance = data.get('itfinance')
    rpo = data.get('rpo')
    poa = data.get('poa')
    cip = data.get('cip')
    mspi = data.get('mspi')
    valmsl = data.get('valmsl')
    valfqmr = data.get('valfqmr')
    valdcap = data.get('valdcap')
    valcifw = data.get('valcifw')
    valoem = data.get('valoem')
    valbnft = data.get('valbnft')
    valnpvv = data.get('valnpvv')
    valacd = data.get('valacd')
    valroi = data.get('valroi')
    valinvestment = data.get('valinvestment')
    valmonths = data.get('valmonths')
    
    if not client_name:
        return "Error: 'client_name' parameter is required", 400





    zip_path = r"template.zip"  # Tam dosya yolunu girin
    output_pptx_path = r"output.pptx"  # Çıkış dosyasının yolunu belirtin

    modify_slide_xml_and_image(zip_path, output_pptx_path,client_name,itfinance,rpo,poa,cip,mspi,valmsl,valfqmr,valdcap,
                               valcifw,valoem,valbnft,valnpvv,valacd,valroi,valinvestment,valmonths)

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
