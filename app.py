from flask import Flask, request, send_file, jsonify,render_template
from pptx import Presentation
from pptx.util import Inches, Pt, Cm
from pptx.dml.color import RGBColor
from bs4 import BeautifulSoup
import os,re
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
    
    donutit = '' if donutit in [None, 0] else donutit
    donutrpo = '' if donutrpo in [None, 0] else donutrpo
    donutpoa = '' if donutpoa in [None, 0] else donutpoa
    donutdcap = '' if donutdcap in [None, 0] else donutdcap
    donutcip = '' if donutcip in [None, 0] else donutcip
    donutmspi = '' if donutmspi in [None, 0] else donutmspi
    donutmsl = '' if donutmsl in [None, 0] else donutmsl
    donutfqmr = '' if donutfqmr in [None, 0] else donutfqmr
    donutcifw = '' if donutcifw in [None, 0] else donutcifw
    donutoem = '' if donutoem in [None, 0] else donutoem

    val = '''<c:numCache>
									<c:formatCode>"£"#,##0</c:formatCode>
									<c:ptCount val="10"/>
									<c:pt idx="0">
										<c:v>{donutrpo}</c:v>
									</c:pt>
									<c:pt idx="1">
										<c:v>{donutpoa}</c:v>
									</c:pt>
									<c:pt idx="2">
										<c:v>{donutcip}</c:v>
									</c:pt>
									<c:pt idx="3">
										<c:v>{donutmspi}</c:v>
									</c:pt>
									<c:pt idx="4">
										<c:v>{donutmsl}</c:v>
									</c:pt>
									<c:pt idx="5">
										<c:v>{donutfqmr}</c:v>
									</c:pt>
									<c:pt idx="6">
										<c:v>{donutdcap}</c:v>
									</c:pt>
									<c:pt idx="7">
										<c:v>{donutcifw}</c:v>
									</c:pt>
									<c:pt idx="8">
										<c:v>{donutit}</c:v>
									</c:pt>
									<c:pt idx="9">
										<c:v>{donutoem}</c:v>
									</c:pt>
								</c:numCache>'''
    with open('donutxml.xml', 'r', encoding='utf-8') as file:
        xml_content = file.read() 

    # Yıl değerlerini xml_content içinde değiştir
    val = val.replace('{donutrpo}', str(donutrpo).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutpoa}', str(donutpoa).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutcip}', str(donutcip).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutmspi}', str(donutmspi).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutmsl}', str(donutmsl).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutfqmr}', str(donutfqmr).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutdcap}', str(donutdcap).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutcifw}', str(donutcifw).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutit}', str(donutit).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutoem}', str(donutoem).replace('£','').replace(',','').replace(' ',''))

    
    return xml_content.replace('numrefval1',val)

def create_clientdonut_xml(donutit,donutrpo,donutpoa,
                             donutdcap,donutcip,donutmspi,donutmsl,donutfqmr,donutcifw,donutoem):
    
    
    # barxml.txt dosyasından XML içeriğini oku
    with open('clientdonut.xml', 'r', encoding='utf-8') as file:
        xml_content = file.read()
    if (str(donutrpo).replace('£','').replace(',','').replace(' ','')) == '0':
        donutrpo=''
    if (str(donutpoa).replace('£','').replace(',','').replace(' ','')) == '0':
        donutpoa=''
    if (str(donutcip).replace('£','').replace(',','').replace(' ','')) == '0':
        donutcip=''
    if (str(donutmspi).replace('£','').replace(',','').replace(' ','')) == '0':
        donutmspi=''
    if (str(donutmsl).replace('£','').replace(',','').replace(' ','')) == '0':
        donutmsl=''
    if (str(donutfqmr).replace('£','').replace(',','').replace(' ','')) == '0':
        donutfqmr=''
    if (str(donutdcap).replace('£','').replace(',','').replace(' ','')) == '0':
        donutdcap=''
    if (str(donutcifw).replace('£','').replace(',','').replace(' ','')) == '0':
        donutcifw=''
    if (str(donutit).replace('£','').replace(',','').replace(' ','')) == '0':
        donutit=''
    if (str(donutoem).replace('£','').replace(',','').replace(' ','')) == '0':
        donutoem=''
    # Yıl değerlerini xml_content içinde değiştir
    xml_content = xml_content.replace('{donutrpo}', str(donutrpo).replace('£','').replace(',','').replace(' ',''))
    xml_content = xml_content.replace('{donutpoa}', str(donutpoa).replace('£','').replace(',','').replace(' ',''))
    xml_content = xml_content.replace('{donutcip}', str(donutcip).replace('£','').replace(',','').replace(' ',''))
    xml_content = xml_content.replace('{donutmspi}', str(donutmspi).replace('£','').replace(',','').replace(' ',''))
    xml_content = xml_content.replace('{donutmsl}', str(donutmsl).replace('£','').replace(',','').replace(' ',''))
    xml_content = xml_content.replace('{donutfqmr}', str(donutfqmr).replace('£','').replace(',','').replace(' ',''))
    xml_content = xml_content.replace('{donutdcap}', str(donutdcap).replace('£','').replace(',','').replace(' ',''))
    xml_content = xml_content.replace('{donutcifw}', str(donutcifw).replace('£','').replace(',','').replace(' ',''))
    xml_content = xml_content.replace('{donutit}', str(donutit).replace('£','').replace(',','').replace(' ',''))
    xml_content = xml_content.replace('{donutoem}', str(donutoem).replace('£','').replace(',','').replace(' ',''))
    
    return xml_content

def create_questionaredonut_xml(donutit,donutrpo,donutpoa,
                             donutdcap,donutcip,donutmspi,donutmsl,donutfqmr,donutcifw,donutoem):
    
    
    # barxml.txt dosyasından XML içeriğini oku
    with open('qdonutxml.xml', 'r', encoding='utf-8') as file:
        xml_content = file.read()
    if (str(donutrpo).replace('£','').replace(',','').replace(' ','')) == '0':
        donutrpo=''
    if (str(donutpoa).replace('£','').replace(',','').replace(' ','')) == '0':
        donutpoa=''
    if (str(donutcip).replace('£','').replace(',','').replace(' ','')) == '0':
        donutcip=''
    if (str(donutmspi).replace('£','').replace(',','').replace(' ','')) == '0':
        donutmspi=''
    if (str(donutmsl).replace('£','').replace(',','').replace(' ','')) == '0':
        donutmsl=''
    if (str(donutfqmr).replace('£','').replace(',','').replace(' ','')) == '0':
        donutfqmr=''
    if (str(donutdcap).replace('£','').replace(',','').replace(' ','')) == '0':
        donutdcap=''
    if (str(donutcifw).replace('£','').replace(',','').replace(' ','')) == '0':
        donutcifw=''
    if (str(donutit).replace('£','').replace(',','').replace(' ','')) == '0':
        donutit=''
    if (str(donutoem).replace('£','').replace(',','').replace(' ','')) == '0':
        donutoem=''
    # Yıl değerlerini xml_content içinde değiştir
    xml_content = xml_content.replace('{donutrpo}', str(donutrpo).replace('£','').replace(',','').replace(' ',''))
    xml_content = xml_content.replace('{donutpoa}', str(donutpoa).replace('£','').replace(',','').replace(' ',''))
    xml_content = xml_content.replace('{donutcip}', str(donutcip).replace('£','').replace(',','').replace(' ',''))
    xml_content = xml_content.replace('{donutmspi}', str(donutmspi).replace('£','').replace(',','').replace(' ',''))
    xml_content = xml_content.replace('{donutmsl}', str(donutmsl).replace('£','').replace(',','').replace(' ',''))
    xml_content = xml_content.replace('{donutfqmr}', str(donutfqmr).replace('£','').replace(',','').replace(' ',''))
    xml_content = xml_content.replace('{donutdcap}', str(donutdcap).replace('£','').replace(',','').replace(' ',''))
    xml_content = xml_content.replace('{donutcifw}', str(donutcifw).replace('£','').replace(',','').replace(' ',''))
    xml_content = xml_content.replace('{donutit}', str(donutit).replace('£','').replace(',','').replace(' ',''))
    xml_content = xml_content.replace('{donutoem}', str(donutoem).replace('£','').replace(',','').replace(' ',''))
    
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

def update_zip_with_new_xml_client(zip_path, output_zip_path, year1invest, year1return, year2invest, year2return,
                             year3invest, year3return, year4invest, year4return, year5invest, year5return,donutit,donutrpo,donutpoa,
                             donutdcap,donutcip,donutmspi,donutmsl,donutfqmr,donutcifw,donutoem):
    temp_dir = 'temp_zip'
    os.makedirs(temp_dir, exist_ok=True)

    # ZIP dosyasını çıkar
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    
    donut_xml_content_client = create_clientdonut_xml(donutit,donutrpo,donutpoa,
                             donutdcap,donutcip,donutmspi,donutmsl,donutfqmr,donutcifw,donutoem)


    new_xml_path = os.path.join(temp_dir, 'ppt', 'charts', 'chart1.xml')
    with open(new_xml_path, 'w', encoding='utf-8') as xml_file:
        xml_file.write(donut_xml_content_client)

    # Güncellenmiş dosyaları yeni ZIP dosyası olarak kaydet
    with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
        for foldername, subfolders, filenames in os.walk(temp_dir):
            for filename in filenames:
                filepath = os.path.join(foldername, filename)
                arcname = os.path.relpath(filepath, temp_dir)
                zip_ref.write(filepath, arcname)

    # Geçici dizini temizle
    shutil.rmtree(temp_dir)


def update_zip_with_new_xml_questionare(zip_path, output_zip_path, baitval,barpoval,bapoaval,bacipval,bamspival,bamslval,bafqmrval,badcapval,
                               bacifwval,baoemval,batotalval,donutit,donutrpo,donutpoa,
                             donutdcap,donutcip,donutmspi,donutmsl,donutfqmr,donutcifw,donutoem):
    temp_dir = 'temp_zip'
    os.makedirs(temp_dir, exist_ok=True)

    # ZIP dosyasını çıkar
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    
    donut_xml_content_client = create_questionaredonut_xml(donutit,donutrpo,donutpoa,
                             donutdcap,donutcip,donutmspi,donutmsl,donutfqmr,donutcifw,donutoem)


    new_xml_path = os.path.join(temp_dir, 'ppt', 'charts', 'chart1.xml')
    with open(new_xml_path, 'w', encoding='utf-8') as xml_file:
        xml_file.write(donut_xml_content_client)

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
                               donutfqmr='0',donutcifw='0',donutoem='0',
                               prpoval='0',ppoaval='0',pcipval='0',pmspival='0',pmslval='0',pfqmrval='0',
                                pdcapval='0',pcifwval='0',poemval='0',pitfinanceval='0',totalcostval=0,per1x='0',per2x='0',
                                per3x='0',per4x='0',per5x='0',per6x='0',per7x='0',per8x='0',per9x='0',per10x='0',asrpoval='0',aspoaval='0',ascipval='0',asmspival='0',asmslval='0',
                                asfqmrval='0',asdcapval='0',ascifwval='0',asoemval='0',asitfinance='0',
                                eytimerpo='0',eytimepoa='0',eytimecip='0',eytimemspi='0',eytimemsl='0',eytimefqmr='0',eytimedcap='0',
                                eytimecifw='0',eytimeoem='0',eytimeitfinance='0',q1itfinance='',q1rpo='',q2rpo='',q3rpo='',q1poa='',q2poa='',q3poa='',q1cip='',
    q2cip='',q3cip='',q4cip='',q1mspi='',q2mspi='',q3mspi='',q1msl='',q2msl='',q1fqmr='',
    q2fqmr='',q3fqmr='',q4fqmr='',q1dcap='',q2dcap='',q3dcap='',q1cifw='',q2cifw='',
    q0itfinance='hosted',q0rpo='this is a gap today',q0poa='this is a gap today',q0mspi='manual entry',q0msl='this is a gap today',q0fqmr='this is a gap today',q0dcap='this is a gap today',q0cifw='this is a gap today',valsubmissionid=''):
    
    
    if q0itfinance==" ":
        q0itfinance='hosted'
    if q0rpo==" ":
        q0rpo='this is a gap today'
    if q0itfinance==" ":
        q0itfinance='hosted'
    if q0poa==" ":
        q0poa='this is a gap today'
    if q0mspi==" ":
        q0mspi='manual entry'
    if q0msl==" ":
        q0msl='this is a gap today'
    if q0fqmr==" ":
        q0fqmr='this is a gap today'
    if q0dcap==" ":
        q0dcap='this is a gap today'
    if q0cifw==" ":
        q0cifw='this is a gap today'

    # Geçici çalışma dizinini oluştur
    temp_dir = 'temp_pptx'
    os.makedirs(temp_dir, exist_ok=True)


    # .pptx dosyasını aç ve dosyaları çıkar
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    # slides klasöründeki tüm slide XML dosyalarını bul ve işle
    slides_dir = os.path.join(temp_dir, 'ppt', 'slides')
    slide_files = [f for f in os.listdir(slides_dir) if f.startswith('slide') and f.endswith('.xml')]

    chart_file = os.path.join(temp_dir, 'ppt', 'charts', 'chart1.xml')
    with open(chart_file, 'r', encoding='utf-8') as file:
        content = file.read()

    color_hosted='ED7D31'
    color_happy='92D050'
    color_better='FFC000'
    color_manual='C00000'
    color_gap='C00000'


    # <c:numRef> etiketleri arasındaki numRefValue'yu değiştir
    if (str(donutrpo).replace('£','').replace(',','').replace(' ','')) == '0':
        donutrpo=''
    if (str(donutpoa).replace('£','').replace(',','').replace(' ','')) == '0':
        donutpoa=''
    if (str(donutcip).replace('£','').replace(',','').replace(' ','')) == '0':
        donutcip=''
    if (str(donutmspi).replace('£','').replace(',','').replace(' ','')) == '0':
        donutmspi=''
    if (str(donutmsl).replace('£','').replace(',','').replace(' ','')) == '0':
        donutmsl=''
    if (str(donutfqmr).replace('£','').replace(',','').replace(' ','')) == '0':
        donutfqmr=''
    if (str(donutdcap).replace('£','').replace(',','').replace(' ','')) == '0':
        donutdcap=''
    if (str(donutcifw).replace('£','').replace(',','').replace(' ','')) == '0':
        donutcifw=''
    if (str(donutit).replace('£','').replace(',','').replace(' ','')) == '0':
        donutit=''
    if (str(donutoem).replace('£','').replace(',','').replace(' ','')) == '0':
        donutoem=''

    val = '''<c:numCache>
									<c:formatCode>"£"#,##0</c:formatCode>
									<c:ptCount val="10"/>
									<c:pt idx="0">
										<c:v>{donutrpo}</c:v>
									</c:pt>
									<c:pt idx="1">
										<c:v>{donutpoa}</c:v>
									</c:pt>
									<c:pt idx="2">
										<c:v>{donutcip}</c:v>
									</c:pt>
									<c:pt idx="3">
										<c:v>{donutmspi}</c:v>
									</c:pt>
									<c:pt idx="4">
										<c:v>{donutmsl}</c:v>
									</c:pt>
									<c:pt idx="5">
										<c:v>{donutfqmr}</c:v>
									</c:pt>
									<c:pt idx="6">
										<c:v>{donutdcap}</c:v>
									</c:pt>
									<c:pt idx="7">
										<c:v>{donutcifw}</c:v>
									</c:pt>
									<c:pt idx="8">
										<c:v>{donutit}</c:v>
									</c:pt>
									<c:pt idx="9">
										<c:v>{donutoem}</c:v>
									</c:pt>
								</c:numCache>'''
    
    # Yıl değerlerini xml_content içinde değiştir
    val = val.replace('{donutrpo}', str(donutrpo).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutpoa}', str(donutpoa).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutcip}', str(donutcip).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutmspi}', str(donutmspi).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutmsl}', str(donutmsl).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutfqmr}', str(donutfqmr).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutdcap}', str(donutdcap).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutcifw}', str(donutcifw).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutit}', str(donutit).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutoem}', str(donutoem).replace('£','').replace(',','').replace(' ',''))

    updated_content = re.sub(r'(<c:numRef>)(.*?)(</c:numRef>)', r'\1\n\t\t\t\t\t\t\t\t\t' + val + r'\n\t\t\t\t\t\t\3', content, flags=re.DOTALL)
    #updated_content = re.sub(r'<c:pt idx="(\d+)">\s*<c:v>0</c:v>\s*</c:pt>', '', content)
    # Güncellenmiş içeriği tekrar dosyaya yaz
    with open(chart_file, 'w', encoding='utf-8') as file:
        file.write(updated_content)
    

    zip_path = 'template.zip'  # Güncellemek istediğin template.zip
    output_zip_path = 'template.zip'  # Çıkış dosyasının adı

    
    for slide_file in slide_files:
        slide_xml_path = os.path.join(slides_dir, slide_file)
        
        # XML dosyasını parse et
        tree = ET.parse(slide_xml_path)
        root = tree.getroot()

        # XML namespace tanımı (değiştirebilir)
        namespace = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
        

    # a:srgbClr etiketlerini ara ve değiştir
        for elem in root.findall('.//a:srgbClr', namespace):  # 'a:srgbClr' tag'ini ara
            if elem.attrib.get('val') == '011111':  # Eğer 'val' özelliği '111111' ise
                if q0itfinance.lower()=='hosted':
                    elem.set('val', color_hosted)
                elif q0itfinance.lower()=='happy as it is':
                    elem.set('val', color_happy)
                elif q0itfinance.lower()=='could be better':
                    elem.set('val', color_better)
                elif q0itfinance.lower()=='manual entry':
                    elem.set('val', color_manual)
                elif q0itfinance.lower()=='this is a gap today':
                    elem.set('val', color_gap)

            if elem.attrib.get('val') == '111111':  # Eğer 'val' özelliği '111111' ise
                if q0rpo.lower()=='hosted':
                    elem.set('val', color_hosted)
                elif q0rpo.lower()=='happy as it is':
                    elem.set('val', color_happy)
                elif q0rpo.lower()=='could be better':
                    elem.set('val', color_better)
                elif q0rpo.lower()=='manual entry':
                    elem.set('val', color_manual)
                elif q0rpo.lower()=='this is a gap today':
                    elem.set('val', color_gap)

            if elem.attrib.get('val') == '211111':  # Eğer 'val' özelliği '111111' ise
                if q0poa.lower()=='hosted':
                    elem.set('val', color_hosted)
                elif q0poa.lower()=='happy as it is':
                    elem.set('val', color_happy)
                elif q0poa.lower()=='could be better':
                    elem.set('val', color_better)
                elif q0poa.lower()=='manual entry':
                    elem.set('val', color_manual)
                elif q0poa.lower()=='this is a gap today':
                    elem.set('val', color_gap)

            if elem.attrib.get('val') == '411111':  # Eğer 'val' özelliği '111111' ise
                if q0mspi.lower()=='hosted':
                    elem.set('val', color_hosted)
                elif q0mspi.lower()=='happy as it is':
                    elem.set('val', color_happy)
                elif q0mspi.lower()=='could be better':
                    elem.set('val', color_better)
                elif q0mspi.lower()=='manual entry':
                    elem.set('val', color_manual)
                elif q0mspi.lower()=='this is a gap today':
                    elem.set('val', color_gap)

            if elem.attrib.get('val') == '511111':  # Eğer 'val' özelliği '111111' ise
                if q0msl.lower()=='hosted':
                    elem.set('val', color_hosted)
                elif q0msl.lower()=='happy as it is':
                    elem.set('val', color_happy)
                elif q0msl.lower()=='could be better':
                    elem.set('val', color_better)
                elif q0msl.lower()=='manual entry':
                    elem.set('val', color_manual)
                elif q0msl.lower()=='this is a gap today':
                    elem.set('val', color_gap)

            if elem.attrib.get('val') == '611111':  # Eğer 'val' özelliği '111111' ise
                if q0fqmr.lower()=='hosted':
                    elem.set('val', color_hosted)
                elif q0fqmr.lower()=='happy as it is':
                    elem.set('val', color_happy)
                elif q0fqmr.lower()=='could be better':
                    elem.set('val', color_better)
                elif q0fqmr.lower()=='manual entry':
                    elem.set('val', color_manual)
                elif q0fqmr.lower()=='this is a gap today':
                    elem.set('val', color_gap)

            if elem.attrib.get('val') == '711111':  # Eğer 'val' özelliği '111111' ise
                if q0dcap.lower()=='hosted':
                    elem.set('val', color_hosted)
                elif q0dcap.lower()=='happy as it is':
                    elem.set('val', color_happy)
                elif q0dcap.lower()=='could be better':
                    elem.set('val', color_better)
                elif q0dcap.lower()=='manual entry':
                    elem.set('val', color_manual)
                elif q0dcap.lower()=='this is a gap today':
                    elem.set('val', color_gap)

            if elem.attrib.get('val') == '811111':  # Eğer 'val' özelliği '111111' ise
                if q0cifw.lower()=='hosted':
                    elem.set('val', color_hosted)
                elif q0cifw.lower()=='happy as it is':
                    elem.set('val', color_happy)
                elif q0cifw.lower()=='could be better':
                    elem.set('val', color_better)
                elif q0cifw.lower()=='manual entry':
                    elem.set('val', color_manual)
                elif q0cifw.lower()=='this is a gap today':
                    elem.set('val', color_gap)
        
        for elem in root.findall('.//a:t', namespace):

            if 'valsubmissionid' in elem.text:
                elem.text = elem.text.replace('valsubmissionid', str(valsubmissionid))
            if 'q0itfinance' in elem.text:
                elem.text = elem.text.replace('q0itfinance', str(q0itfinance))

            if 'q0rpo' in elem.text:
                elem.text = elem.text.replace('q0rpo', str(q0rpo))

            if 'q0poa' in elem.text:
                elem.text = elem.text.replace('q0poa', str(q0poa))

            if 'q0mspi' in elem.text:
                elem.text = elem.text.replace('q0mspi', str(q0mspi))
            
            if 'q0msl' in elem.text:
                elem.text = elem.text.replace('q0msl', str(q0msl))

            if 'q0fqmr' in elem.text:
                elem.text = elem.text.replace('q0fqmr', str(q0fqmr))

            if 'q0dcap' in elem.text:
                elem.text = elem.text.replace('q0dcap', str(q0dcap))

            if 'q0cifw' in elem.text:
                elem.text = elem.text.replace('q0cifw', str(q0cifw))




            if 'q1itfinance' in elem.text:
                elem.text = elem.text.replace('q1itfinance', str(q1itfinance))

            if 'q1rpo' in elem.text:
                elem.text = elem.text.replace('q1rpo', str(q1rpo))
            
            if 'q2rpo' in elem.text:
                elem.text = elem.text.replace('q2rpo', str(q2rpo))
            
            if 'q3rpo' in elem.text:
                elem.text = elem.text.replace('q3rpo', str(q3rpo))

            if 'q1poa' in elem.text:
                elem.text = elem.text.replace('q1poa', str(q1poa))

            if 'q2poa' in elem.text:
                elem.text = elem.text.replace('q2poa', str(q2poa))

            if 'q3poa' in elem.text:
                elem.text = elem.text.replace('q3poa', str(q3poa))

            if 'q1cip' in elem.text:
                elem.text = elem.text.replace('q1cip', str(q1cip))

            if 'q2cip' in elem.text:
                elem.text = elem.text.replace('q2cip', str(q2cip))

            if 'q3cip' in elem.text:
                elem.text = elem.text.replace('q3cip', str(q3cip))

            if 'q4cip' in elem.text:
                elem.text = elem.text.replace('q4cip', str(q4cip))
            
            if 'q1mspi' in elem.text:
                elem.text = elem.text.replace('q1mspi', str(q1mspi))
            
            if 'q2mspi' in elem.text:
                elem.text = elem.text.replace('q2mspi', str(q2mspi))
            
            if 'q3mspi' in elem.text:
                elem.text = elem.text.replace('q3mspi', str(q3mspi))
            
            if 'q1msl' in elem.text:
                elem.text = elem.text.replace('q1msl', str(q1msl))

            if 'q2msl' in elem.text:
                elem.text = elem.text.replace('q2msl', str(q2msl))

            if 'q1fqmr' in elem.text:
                elem.text = elem.text.replace('q1fqmr', str(q1fqmr))

            if 'q2fqmr' in elem.text:
                elem.text = elem.text.replace('q2fqmr', str(q2fqmr))

            if 'q3fqmr' in elem.text:
                elem.text = elem.text.replace('q3fqmr', str(q3fqmr))

            if 'q4fqmr' in elem.text:
                elem.text = elem.text.replace('q4fqmr', str(q4fqmr))

            if 'q1dcap' in elem.text:
                elem.text = elem.text.replace('q1dcap', str(q1dcap))

            if 'q2dcap' in elem.text:
                elem.text = elem.text.replace('q2dcap', str(q2dcap))

            if 'q3dcap' in elem.text:
                elem.text = elem.text.replace('q3dcap', str(q3dcap))

            if 'q1cifw' in elem.text:
                elem.text = elem.text.replace('q1cifw', str(q1cifw))

            if 'q2cifw' in elem.text:
                elem.text = elem.text.replace('q2cifw', str(q2cifw))


            if 'eytimerpo' in elem.text:
                elem.text = elem.text.replace('eytimerpo', str(eytimerpo))
            if 'eytimepoa' in elem.text:
                elem.text = elem.text.replace('eytimepoa', str(eytimepoa))
            if 'eytimecip' in elem.text:
                elem.text = elem.text.replace('eytimecip', str(eytimecip))
            if 'eytimemspi' in elem.text:
                elem.text = elem.text.replace('eytimemspi', str(eytimemspi))
            if 'eytimemsl' in elem.text:
                elem.text = elem.text.replace('eytimemsl', str(eytimemsl))
            if 'eytimefqmr' in elem.text:
                elem.text = elem.text.replace('eytimefqmr', str(eytimefqmr))
            if 'eytimedcap' in elem.text:
                elem.text = elem.text.replace('eytimedcap', str(eytimedcap))
            if 'eytimecifw' in elem.text:
                elem.text = elem.text.replace('eytimecifw', str(eytimecifw))
            if 'eytimeoem' in elem.text:
                elem.text = elem.text.replace('eytimeoem', str(eytimeoem))
            if 'eytimeitfinance' in elem.text:
                elem.text = elem.text.replace('eytimeitfinance', str(eytimeitfinance))


            if 'asrpoval' in elem.text:
                elem.text = elem.text.replace('asrpoval', str(format_with_commas(asrpoval)).replace(' ',''))
            if 'aspoaval' in elem.text:
                elem.text = elem.text.replace('aspoaval', str(format_with_commas(aspoaval)).replace(' ',''))
            if 'ascipval' in elem.text:
                elem.text = elem.text.replace('ascipval', str(format_with_commas(ascipval)).replace(' ',''))
            if 'asmspival' in elem.text:
                elem.text = elem.text.replace('asmspival', str(format_with_commas(asmspival)).replace(' ',''))
            if 'asmslval' in elem.text:
                elem.text = elem.text.replace('asmslval', str(format_with_commas(asmslval)).replace(' ',''))
            if 'asfqmrval' in elem.text:
                elem.text = elem.text.replace('asfqmrval', str(format_with_commas(asfqmrval)).replace(' ',''))
            if 'asdcapval' in elem.text:
                elem.text = elem.text.replace('asdcapval', str(format_with_commas(asdcapval)).replace(' ',''))
            if 'ascifwval' in elem.text:
                elem.text = elem.text.replace('ascifwval', str(format_with_commas(ascifwval)).replace(' ',''))
            if 'asoemval' in elem.text:
                elem.text = elem.text.replace('asoemval', str(format_with_commas(asoemval)).replace(' ',''))
            if 'asitfinance' in elem.text:
                elem.text = elem.text.replace('asitfinance', str(format_with_commas(asitfinance)).replace(' ',''))

            if 'rpoper' in elem.text:
                elem.text = elem.text.replace('rpoper', str(rpoper))
            if 'itfinanceper' in elem.text:
                elem.text = elem.text.replace('itfinanceper', str(itfinanceper))
            if 'poaper' in elem.text:
                elem.text = elem.text.replace('poaper', str(poaper))
            if 'cipper' in elem.text:
                elem.text = elem.text.replace('cipper', str(cipper))
            if 'mspiper' in elem.text:
                elem.text = elem.text.replace('mspiper', str(mspiper))
            if 'valmslper' in elem.text:
                elem.text = elem.text.replace('valmslper', str(valmslper))
            if 'valfqmrper' in elem.text:
                elem.text = elem.text.replace('valfqmrper', str(valfqmrper))
            if 'valdcapper' in elem.text:
                elem.text = elem.text.replace('valdcapper', str(valdcapper))
            if 'valcifwper' in elem.text:
                elem.text = elem.text.replace('valcifwper', str(valcifwper))

            if 'per1x' in elem.text:
                elem.text = elem.text.replace('per1x', str(per1x))
            if 'per2x' in elem.text:
                elem.text = elem.text.replace('per2x', str(per2x))
            if 'per3x' in elem.text:
                elem.text = elem.text.replace('per3x', str(per3x))
            if 'per4x' in elem.text:
                elem.text = elem.text.replace('per4x', str(per4x))
            if 'per5x' in elem.text:
                elem.text = elem.text.replace('per5x', str(per5x))
            if 'per6x' in elem.text:
                elem.text = elem.text.replace('per6x', str(per6x))
            if 'per7x' in elem.text:
                elem.text = elem.text.replace('per7x', str(per7x))
            if 'per8x' in elem.text:
                elem.text = elem.text.replace('per8x', str(per8x))
            if 'per9x' in elem.text:
                elem.text = elem.text.replace('per9x', str(per9x))
            if 'per10x' in elem.text:
                elem.text = elem.text.replace('per10x', str(per10x))

            if 'prpoval' in elem.text:
                elem.text = elem.text.replace('prpoval', prpoval.replace('£','').replace(' ',''))
            if 'ppoaval' in elem.text:
                elem.text = elem.text.replace('ppoaval', ppoaval.replace('£','').replace(' ',''))
            if 'pcipval' in elem.text:
                elem.text = elem.text.replace('pcipval', pcipval.replace('£','').replace(' ',''))
            if 'pmspival' in elem.text:
                elem.text = elem.text.replace('pmspival', pmspival.replace('£','').replace(' ',''))
            if 'pmslval' in elem.text:
                elem.text = elem.text.replace('pmslval', pmslval.replace('£','').replace(' ',''))
            if 'pfqmrval' in elem.text:
                elem.text = elem.text.replace('pfqmrval', pfqmrval.replace('£','').replace(' ',''))
            if 'pdcapval' in elem.text:
                elem.text = elem.text.replace('pdcapval', pdcapval.replace('£','').replace(' ',''))
            if 'pcifwval' in elem.text:
                elem.text = elem.text.replace('pcifwval', pcifwval.replace('£','').replace(' ',''))
            if 'poemval' in elem.text:
                elem.text = elem.text.replace('poemval', poemval.replace('£','').replace(' ',''))
            if 'pitfinanceval' in elem.text:
                elem.text = elem.text.replace('pitfinanceval', pitfinanceval.replace('£',''))
        # XML içeriğinde £XX,000 ifadelerini sırayla değiştir
        for elem in root.findall('.//a:t', namespace):
            if 'valclient' in elem.text:
                elem.text = elem.text.replace('valclient', client_name)
            if 'itfinance' in elem.text:
                elem.text = elem.text.replace('itfinance', itfinance.replace('£','').replace(' ',''))
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
                elem.text = elem.text.replace('valcostof', costofdoingnothing1.replace('£','').replace(' ',''))            
            if 'totalcostval' in elem.text:
                elem.text = elem.text.replace('totalcostval',str(format_with_commas(totalcostval)))
        

        # Yeni XML dosyasını oluştur ve ZIP dosyasını güncelle
        

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

def modify_slide_xml_and_image_client(zip_path, output_pptx_path,client_name,
                               itfinance='0',rpo='0',poa='0',cip='0',mspi='0',valmsl='0',valfqmr='0',valdcap='0',
                               valcifw='0',valoem='0',valbnft='0',valnpvv='0',valacd='0',valroi='0',valinvestment='0',valmonths='0',valhours='0',
                               year1invest='0',
                                year1return='0',year2invest='0',year2return='0',year3invest='0',
                                year3return='0',year4invest='0',year4return='0',year5invest='0',
                                year5return='0',costofdoingnothing1='0',itfinanceper='0',rpoper='0',poaper='0',cipper='0'
                                ,mspiper='0',valmslper='0',valfqmrper='0',valdcapper='0',
                               valcifwper='0',valoemper='0',
                               donutit='0',donutrpo='0',donutpoa='0',donutdcap='0',donutcip='0',donutmspi='0',donutmsl='0',
                               donutfqmr='0',donutcifw='0',donutoem='0',
                               prpoval='0',ppoaval='0',pcipval='0',pmspival='0',pmslval='0',pfqmrval='0',
                                pdcapval='0',pcifwval='0',poemval='0',pitfinanceval='0',totalcostval='0',per1x='0',per2x='0',
                                per3x='0',per4x='0',per5x='0',per6x='0',per7x='0',per8x='0',per9x='0',per10x='0',
                                valsubmissionid=' ',q0mspi=' ',q1mspi=' ',q2mspi=' ',q3mspi=' ',q0fqmr=' ',q1fqmr=' ',
    q2fqmr=' ',q3fqmr=' ',q4fqmr=' ',q0dcap=' ',q1dcap=' ',q2dcap=' ',q3dcap=' '):

  
    
    # Geçici çalışma dizinini oluştur
    temp_dir = 'temp_pptx'
    os.makedirs(temp_dir, exist_ok=True)

    color_hosted='ED7D31'
    color_happy='92D050'
    color_better='FFC000'
    color_manual='C00000'
    color_gap='C00000'

    # .pptx dosyasını aç ve dosyaları çıkar
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    chart_file = os.path.join(temp_dir, 'ppt', 'charts', 'chart1.xml')
    with open(chart_file, 'r', encoding='utf-8') as file:
        content = file.read()


     

    # <c:numRef> etiketleri arasındaki numRefValue'yu değiştir
    if (str(donutrpo).replace('£','').replace(',','').replace(' ','')) == '0':
        donutrpo=''
    if (str(donutpoa).replace('£','').replace(',','').replace(' ','')) == '0':
        donutpoa=''
    if (str(donutcip).replace('£','').replace(',','').replace(' ','')) == '0':
        donutcip=''
    if (str(donutmspi).replace('£','').replace(',','').replace(' ','')) == '0':
        donutmspi=''
    if (str(donutmsl).replace('£','').replace(',','').replace(' ','')) == '0':
        donutmsl=''
    if (str(donutfqmr).replace('£','').replace(',','').replace(' ','')) == '0':
        donutfqmr=''
    if (str(donutdcap).replace('£','').replace(',','').replace(' ','')) == '0':
        donutdcap=''
    if (str(donutcifw).replace('£','').replace(',','').replace(' ','')) == '0':
        donutcifw=''
    if (str(donutit).replace('£','').replace(',','').replace(' ','')) == '0':
        donutit=''
    if (str(donutoem).replace('£','').replace(',','').replace(' ','')) == '0':
        donutoem=''

    val = '''<c:numCache>
									<c:formatCode>"£"#,##0</c:formatCode>
									<c:ptCount val="10"/>
									<c:pt idx="0">
										<c:v>{donutrpo}</c:v>
									</c:pt>
									<c:pt idx="1">
										<c:v>{donutpoa}</c:v>
									</c:pt>
									<c:pt idx="2">
										<c:v>{donutcip}</c:v>
									</c:pt>
									<c:pt idx="3">
										<c:v>{donutmspi}</c:v>
									</c:pt>
									<c:pt idx="4">
										<c:v>{donutmsl}</c:v>
									</c:pt>
									<c:pt idx="5">
										<c:v>{donutfqmr}</c:v>
									</c:pt>
									<c:pt idx="6">
										<c:v>{donutdcap}</c:v>
									</c:pt>
									<c:pt idx="7">
										<c:v>{donutcifw}</c:v>
									</c:pt>
									<c:pt idx="8">
										<c:v>{donutit}</c:v>
									</c:pt>
									<c:pt idx="9">
										<c:v>{donutoem}</c:v>
									</c:pt>
								</c:numCache>'''
    
    # Yıl değerlerini xml_content içinde değiştir
    val = val.replace('{donutrpo}', str(donutrpo).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutpoa}', str(donutpoa).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutcip}', str(donutcip).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutmspi}', str(donutmspi).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutmsl}', str(donutmsl).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutfqmr}', str(donutfqmr).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutdcap}', str(donutdcap).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutcifw}', str(donutcifw).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutit}', str(donutit).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutoem}', str(donutoem).replace('£','').replace(',','').replace(' ',''))

    updated_content = re.sub(r'(<c:numRef>)(.*?)(</c:numRef>)', r'\1\n\t\t\t\t\t\t\t\t\t' + val + r'\n\t\t\t\t\t\t\3', content, flags=re.DOTALL)
    #updated_content = re.sub(r'<c:pt idx="(\d+)">\s*<c:v>0</c:v>\s*</c:pt>', '', content)
    # Güncellenmiş içeriği tekrar dosyaya yaz
    with open(chart_file, 'w', encoding='utf-8') as file:
        file.write(updated_content)

    # slides klasöründeki tüm slide XML dosyalarını bul ve işle
    slides_dir = os.path.join(temp_dir, 'ppt', 'slides')
    slide_files = [f for f in os.listdir(slides_dir) if f.startswith('slide') and f.endswith('.xml')]
    zip_path = 'client_template.zip'  # Güncellemek istediğin template.zip
    output_zip_path = 'client_template.zip'  # Çıkış dosyasının adı

        # Yeni XML dosyasını oluştur ve ZIP dosyasını güncelle
    # update_zip_with_new_xml_client(zip_path, output_zip_path,year1invest=year1invest,
    #                             year1return=year1return,year2invest=year2invest,year2return=year2return,year3invest=year3invest,
    #                             year3return=year3return,year4invest=year4invest,year4return=year4return,year5invest=year5invest,
    #                             year5return=year5return,donutit=donutit,donutrpo=donutrpo,donutpoa=donutpoa,donutdcap=donutdcap,
    #                             donutcip=donutcip,donutmspi=donutmspi,donutmsl=donutmsl,donutfqmr=donutfqmr,donutcifw=donutcifw,
    #                             donutoem=donutoem)

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
        
        
        for elem in root.findall('.//a:srgbClr', namespace):  # 'a:srgbClr' tag'ini ara
         
           
           
            if elem.attrib.get('val') == '411111':  # Eğer 'val' özelliği '111111' ise
                if q0mspi.lower()=='hosted':
                    elem.set('val', color_hosted)
                elif q0mspi.lower()=='happy as it is':
                    elem.set('val', color_happy)
                elif q0mspi.lower()=='could be better':
                    elem.set('val', color_better)
                elif q0mspi.lower()=='manual entry':
                    elem.set('val', color_manual)
                elif q0mspi.lower()=='this is a gap today':
                    elem.set('val', color_gap)

          
            if elem.attrib.get('val') == '611111':  # Eğer 'val' özelliği '111111' ise
                if q0fqmr.lower()=='hosted':
                    elem.set('val', color_hosted)
                elif q0fqmr.lower()=='happy as it is':
                    elem.set('val', color_happy)
                elif q0fqmr.lower()=='could be better':
                    elem.set('val', color_better)
                elif q0fqmr.lower()=='manual entry':
                    elem.set('val', color_manual)
                elif q0fqmr.lower()=='this is a gap today':
                    elem.set('val', color_gap)

            if elem.attrib.get('val') == '711111':  # Eğer 'val' özelliği '111111' ise
                if q0dcap.lower()=='hosted':
                    elem.set('val', color_hosted)
                elif q0dcap.lower()=='happy as it is':
                    elem.set('val', color_happy)
                elif q0dcap.lower()=='could be better':
                    elem.set('val', color_better)
                elif q0dcap.lower()=='manual entry':
                    elem.set('val', color_manual)
                elif q0dcap.lower()=='this is a gap today':
                    elem.set('val', color_gap)

        for elem in root.findall('.//a:t', namespace):
            if 'valsubmissionid' in elem.text:
                elem.text = elem.text.replace('valsubmissionid', str(valsubmissionid))

            if 'q0mspi' in elem.text:
                elem.text = elem.text.replace('q0mspi', str(q0mspi))

            if 'q1mspi' in elem.text:
                elem.text = elem.text.replace('q1mspi', str(q1mspi))

            if 'q2mspi' in elem.text:
                elem.text = elem.text.replace('q2mspi', str(q2mspi))

            if 'q3mspi' in elem.text:
                elem.text = elem.text.replace('q3mspi', str(q3mspi))

            if 'q0fqmr' in elem.text:
                elem.text = elem.text.replace('q0fqmr', str(q0fqmr))

            if 'q1fqmr' in elem.text:
                elem.text = elem.text.replace('q1fqmr', str(q1fqmr))

            if 'q2fqmr' in elem.text:
                elem.text = elem.text.replace('q2fqmr', str(q2fqmr))

            if 'q3fqmr' in elem.text:
                elem.text = elem.text.replace('q3fqmr', str(q3fqmr))

            if 'q4fqmr' in elem.text:
                elem.text = elem.text.replace('q4fqmr', str(q4fqmr))

            if 'q0dcap' in elem.text:
                elem.text = elem.text.replace('q0dcap', str(q0dcap))

            if 'q1dcap' in elem.text:
                elem.text = elem.text.replace('q1dcap', str(q1dcap))

            if 'q2dcap' in elem.text:
                elem.text = elem.text.replace('q2dcap', str(q2dcap))

            if 'q3dcap' in elem.text:
                elem.text = elem.text.replace('q3dcap', str(q3dcap))





            if 'per1x' in elem.text:
                elem.text = elem.text.replace('per1x', str(per1x))
            if 'per2x' in elem.text:
                elem.text = elem.text.replace('per2x', str(per2x))
            if 'per3x' in elem.text:
                elem.text = elem.text.replace('per3x', str(per3x))
            if 'per4x' in elem.text:
                elem.text = elem.text.replace('per4x', str(per4x))
            if 'per5x' in elem.text:
                elem.text = elem.text.replace('per5x', str(per5x))
            if 'per6x' in elem.text:
                elem.text = elem.text.replace('per6x', str(per6x))
            if 'per7x' in elem.text:
                elem.text = elem.text.replace('per7x', str(per7x))
            if 'per8x' in elem.text:
                elem.text = elem.text.replace('per8x', str(per8x))
            if 'per9x' in elem.text:
                elem.text = elem.text.replace('per9x', str(per9x))
            if 'per10x' in elem.text:
                elem.text = elem.text.replace('per10x', str(per10x))

            if 'prpoval' in elem.text:
                elem.text = elem.text.replace('prpoval', prpoval.replace('£','').replace(' ',''))
            if 'ppoaval' in elem.text:
                elem.text = elem.text.replace('ppoaval', ppoaval.replace('£','').replace(' ',''))
            if 'pcipval' in elem.text:
                elem.text = elem.text.replace('pcipval', pcipval.replace('£','').replace(' ',''))
            if 'pmspival' in elem.text:
                elem.text = elem.text.replace('pmspival', pmspival.replace('£','').replace(' ',''))
            if 'pmslval' in elem.text:
                elem.text = elem.text.replace('pmslval', pmslval.replace('£','').replace(' ',''))
            if 'pfqmrval' in elem.text:
                elem.text = elem.text.replace('pfqmrval', pfqmrval.replace('£','').replace(' ',''))
            if 'pdcapval' in elem.text:
                elem.text = elem.text.replace('pdcapval', pdcapval.replace('£','').replace(' ',''))
            if 'pcifwval' in elem.text:
                elem.text = elem.text.replace('pcifwval', pcifwval.replace('£','').replace(' ',''))
            if 'poemval' in elem.text:
                elem.text = elem.text.replace('poemval', poemval.replace('£','').replace(' ',''))
            if 'pitfinanceval' in elem.text:
                elem.text = elem.text.replace('pitfinanceval', pitfinanceval.replace('£','').replace(' ',''))
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
            if 'valmspi' in elem.text:
                elem.text = elem.text.replace('valmspi', format_with_commas(mspi.replace('£','')))
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
                elem.text = elem.text.replace('valcostof', costofdoingnothing1.replace('£','').replace(' ',''))
            if 'valdonutpercentvalues' in elem.text:
                elem.text = elem.text.replace('valdonutpercentvalues',donutpercentvals)
            if 'totalcostval' in elem.text:
                elem.text = elem.text.replace('totalcostval',str(format_with_commas(totalcostval)).replace(' ',''))
        

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

def modify_slide_xml_and_image_questionare(zip_path, output_pptx_path,client_name,
                               baitval='0',barpoval='0',bapoaval='0',bacipval='0',bamspival='0',bamslval='0',bafqmrval='0',badcapval='0',
                               bacifwval='0',baoemval='0',batotalval='0',
                               donutit='0',donutrpo='0',donutpoa='0',donutdcap='0',donutcip='0',donutmspi='0',donutmsl='0',
                               donutfqmr='0',donutcifw='0',donutoem='0',prpoval='0',ppoaval='0',pcipval='0',pmspival='0',pmslval='0',
                               pfqmrval='0',pdcapval='0',pcifwval='0',poemval='0',pitfinanceval='0',period1=' ',period2=' '):
    


    # Geçici çalışma dizinini oluştur
    temp_dir = 'questionare_temp_pptx'
    os.makedirs(temp_dir, exist_ok=True)


    # .pptx dosyasını aç ve dosyaları çıkar
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    # slides klasöründeki tüm slide XML dosyalarını bul ve işle
    slides_dir = os.path.join(temp_dir, 'ppt', 'slides')
    slide_files = [f for f in os.listdir(slides_dir) if f.startswith('slide') and f.endswith('.xml')]

    chart_file = os.path.join(temp_dir, 'ppt', 'charts', 'chart1.xml')
    with open(chart_file, 'r', encoding='utf-8') as file:
        content = file.read()
    updated_content = re.sub(r'(<c:numLit>)(.*?)(</c:numLit>)', r'\1\n\t\t\t\t\t\t\t\t\n\t\t\t\t\t\t\3', content, flags=re.DOTALL)
    # <c:numRef> etiketleri arasındaki numRefValue'yu değiştir
    if (str(donutrpo).replace('£','').replace(',','').replace(' ','')) == '0':
        donutrpo=''
    if (str(donutpoa).replace('£','').replace(',','').replace(' ','')) == '0':
        donutpoa=''
    if (str(donutcip).replace('£','').replace(',','').replace(' ','')) == '0':
        donutcip=''
    if (str(donutmspi).replace('£','').replace(',','').replace(' ','')) == '0':
        donutmspi=''
    if (str(donutmsl).replace('£','').replace(',','').replace(' ','')) == '0':
        donutmsl=''
    if (str(donutfqmr).replace('£','').replace(',','').replace(' ','')) == '0':
        donutfqmr=''
    if (str(donutdcap).replace('£','').replace(',','').replace(' ','')) == '0':
        donutdcap=''
    if (str(donutcifw).replace('£','').replace(',','').replace(' ','')) == '0':
        donutcifw=''
    if (str(donutit).replace('£','').replace(',','').replace(' ','')) == '0':
        donutit=''
    if (str(donutoem).replace('£','').replace(',','').replace(' ','')) == '0':
        donutoem=''

    val = '''
									<c:formatCode>"£"#,##0</c:formatCode>
									<c:ptCount val="10"/>
									<c:pt idx="0">
										<c:v>{donutrpo}</c:v>
									</c:pt>
									<c:pt idx="1">
										<c:v>{donutpoa}</c:v>
									</c:pt>
									<c:pt idx="2">
										<c:v>{donutcip}</c:v>
									</c:pt>
									<c:pt idx="3">
										<c:v>{donutmspi}</c:v>
									</c:pt>
									<c:pt idx="4">
										<c:v>{donutmsl}</c:v>
									</c:pt>
									<c:pt idx="5">
										<c:v>{donutfqmr}</c:v>
									</c:pt>
									<c:pt idx="6">
										<c:v>{donutdcap}</c:v>
									</c:pt>
									<c:pt idx="7">
										<c:v>{donutcifw}</c:v>
									</c:pt>
									<c:pt idx="8">
										<c:v>{donutit}</c:v>
									</c:pt>
									<c:pt idx="9">
										<c:v>{donutoem}</c:v>
									</c:pt>
								'''
    
    # Yıl değerlerini xml_content içinde değiştir
    val = val.replace('{donutrpo}', str(donutrpo).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutpoa}', str(donutpoa).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutcip}', str(donutcip).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutmspi}', str(donutmspi).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutmsl}', str(donutmsl).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutfqmr}', str(donutfqmr).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutdcap}', str(donutdcap).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutcifw}', str(donutcifw).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutit}', str(donutit).replace('£','').replace(',','').replace(' ',''))
    val = val.replace('{donutoem}', str(donutoem).replace('£','').replace(',','').replace(' ',''))
    updated_content = re.sub(r'(<c:numLit>)(.*?)(</c:numLit>)', r'\1\n\t\t\t\t\t\t\t\t\t' + val + r'\n\t\t\t\t\t\t\3', content, flags=re.DOTALL)
    #updated_content = re.sub(r'<c:pt idx="(\d+)">\s*<c:v>0</c:v>\s*</c:pt>', '', content)
    # Güncellenmiş içeriği tekrar dosyaya yaz
    print(val)
    with open(chart_file, 'w', encoding='utf-8') as file:
        file.write(updated_content)
    
    zip_path = 'questionare_template.zip'  # Güncellemek istediğin template.zip
    output_zip_path = 'questionare_template.zip'  # Çıkış dosyasının adı

        # Yeni XML dosyasını oluştur ve ZIP dosyasını güncelle
    # update_zip_with_new_xml_questionare(zip_path, output_zip_path,baitval=baitval,barpoval=barpoval,bapoaval=bapoaval,bacipval=bacipval,bamspival=bamspival,bamslval=bamslval,bafqmrval=bafqmrval,badcapval=badcapval,
    #                            bacifwval=bacifwval,baoemval=baoemval,batotalval=batotalval,donutit=donutit,donutrpo=donutrpo,donutpoa=donutpoa,donutdcap=donutdcap,
    #                             donutcip=donutcip,donutmspi=donutmspi,donutmsl=donutmsl,donutfqmr=donutfqmr,donutcifw=donutcifw,
    #                             donutoem=donutoem)
    # Her bir slide dosyasını işle
    for slide_file in slide_files:
        slide_xml_path = os.path.join(slides_dir, slide_file)
        
        # XML dosyasını parse et
        tree = ET.parse(slide_xml_path)
        root = tree.getroot()

        # XML namespace tanımı (değiştirebilir)
        namespace = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}

         
        for elem in root.findall('.//a:t', namespace):
            if 'valclient' in elem.text:
                elem.text = elem.text.replace('valclient', client_name)
            if 'baitval' in elem.text:
                elem.text = elem.text.replace('baitval', format_with_commas(baitval).replace('£',''))
            if 'barpoval' in elem.text:
                elem.text = elem.text.replace('barpoval', format_with_commas(barpoval).replace('£',''))
            if 'bapoaval' in elem.text:
                elem.text = elem.text.replace('bapoaval', bapoaval.replace('£','').replace(' ',''))
            if 'bacipval' in elem.text:
                elem.text = elem.text.replace('bacipval', bacipval.replace('£','').replace(' ',''))
            if 'bamspival' in elem.text:
                elem.text = elem.text.replace('bamspival', bamspival.replace('£','').replace(' ',''))
            if 'bamslval' in elem.text:
                elem.text = elem.text.replace('bamslval', bamslval.replace('£','').replace(' ',''))
            if 'bafqmrval' in elem.text:
                elem.text = elem.text.replace('bafqmrval', bafqmrval.replace('£','').replace(' ',''))
            if 'badcapval' in elem.text:
                elem.text = elem.text.replace('badcapval', badcapval.replace('£','').replace(' ',''))
            if 'bacifwval' in elem.text:
                elem.text = elem.text.replace('bacifwval', bacifwval.replace('£','').replace(' ',''))
            if 'baoemval' in elem.text:
                elem.text = elem.text.replace('baoemval', baoemval.replace('£','').replace(' ',''))
            if 'batotalval' in elem.text:
                elem.text = elem.text.replace('batotalval', batotalval.replace('£','').replace(' ',''))


            if 'prpoval' in elem.text:
                elem.text = elem.text.replace('prpoval', format_with_commas(prpoval.replace('%','')))
            if 'ppoaval' in elem.text:
                elem.text = elem.text.replace('ppoaval', format_with_commas(ppoaval.replace('%','')))
            if 'pcipval' in elem.text:
                elem.text = elem.text.replace('pcipval', format_with_commas(pcipval.replace('%','')))
            if 'pmspival' in elem.text:
                elem.text = elem.text.replace('pmspival', format_with_commas(pmspival.replace('%','')))
            if 'pmslval' in elem.text:
                elem.text = elem.text.replace('pmslval', format_with_commas(pmslval.replace('%','')))
            if 'pfqmrval' in elem.text:
                elem.text = elem.text.replace('pfqmrval', format_with_commas(pfqmrval.replace('%','')))
            if 'pdcapval' in elem.text:
                elem.text = elem.text.replace('pdcapval', format_with_commas(pdcapval.replace('%','')))
            if 'pcifwval' in elem.text:
                elem.text = elem.text.replace('pcifwval', format_with_commas(pcifwval.replace('%','')))
            if 'poemval' in elem.text:
                elem.text = elem.text.replace('poemval', format_with_commas(poemval.replace('%','')))
            if 'pitfinanceval' in elem.text:
                elem.text = elem.text.replace('pitfinanceval', format_with_commas(pitfinanceval.replace('%','')))

            if 'period1val' in elem.text:
                elem.text = elem.text.replace('period1val', period1)
            if 'period2val' in elem.text:
                elem.text = elem.text.replace('period2val', period2)
        

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
    newmaskedval = data.get('newmaskedval') or ""
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

    itfinanceper = data.get('itfinanceper') or "0"
    rpoper = data.get('rpoper') or "0"
    poaper = data.get('poaper') or "0"
    cipper = data.get('cipper') or "0"
    mspiper = data.get('mspiper') or "0"
    valmslper = data.get('valmslper') or "0"
    valfqmrper = data.get('valfqmrper') or "0"
    valdcapper = data.get('valdcapper') or "0"
    valcifwper = data.get('valcifwper') or "0"
    valoemper = data.get('valoemper') or "0"

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

    prpoval = data.get('prpoval') or "0"
    ppoaval = data.get('ppoaval') or "0"
    pcipval = data.get('pcipval') or "0"
    pmspival = data.get('pmspival') or "0"
    pmslval = data.get('pmslval') or "0"
    pfqmrval = data.get('pfqmrval') or "0"
    pdcapval = data.get('pdcapval') or "0"
    pcifwval = data.get('pcifwval') or "0"
    poemval = data.get('poemval') or "0"
    pitfinanceval = data.get('pitfinanceval') or "0"

    asrpoval = data.get('asrpoval') or "0"
    aspoaval = data.get('aspoaval') or "0"
    ascipval = data.get('ascipval') or "0"
    asmspival = data.get('asmspival') or "0"
    asmslval = data.get('asmslval') or "0"
    asfqmrval = data.get('asfqmrval') or "0"
    asdcapval = data.get('asdcapval') or "0"
    ascifwval = data.get('ascifwval') or "0"
    asoemval = data.get('asoemval') or "0"
    asitfinance = data.get('asitfinance') or "0"

    
    eytimerpo = data.get('eytimerpo') or "0"
    eytimepoa = data.get('eytimepoa') or "0"
    eytimecip = data.get('eytimecip') or "0"
    eytimemspi = data.get('eytimemspi') or "0"
    eytimemsl = data.get('eytimemsl') or "0"
    eytimefqmr = data.get('eytimefqmr') or "0"
    eytimedcap = data.get('eytimedcap') or "0"
    eytimecifw = data.get('eytimecifw') or "0"
    eytimeoem = data.get('eytimeoem') or "0"
    eytimeitfinance = data.get('eytimeitfinance') or "0"

    q1itfinance = data.get('q1itfinance') or " "
    q1rpo = data.get('q1rpo') or " "
    q2rpo = data.get('q2rpo') or " "
    q3rpo = data.get('q3rpo') or " "
    q1poa = data.get('q1poa') or " "
    q2poa = data.get('q2poa') or " "
    q3poa = data.get('q3poa') or " "
    q1cip = data.get('q1cip') or " "
    q2cip = data.get('q2cip') or " "
    q3cip = data.get('q3cip') or " "
    q4cip = data.get('q4cip') or " "
    q1mspi = data.get('q1mspi') or " "
    q2mspi = data.get('q2mspi') or " "
    q3mspi = data.get('q3mspi') or " "
    q1msl = data.get('q1msl') or " "
    q2msl = data.get('q2msl') or " "
    q1fqmr = data.get('q1fqmr') or " "
    q2fqmr = data.get('q2fqmr') or " "
    q3fqmr = data.get('q3fqmr') or " "
    q4fqmr = data.get('q4fqmr') or " "
    q1dcap = data.get('q1dcap') or " "
    q2dcap = data.get('q2dcap') or " "
    q3dcap = data.get('q3dcap') or " "
    q1cifw = data.get('q1cifw') or " "
    q2cifw = data.get('q2cifw') or " "

    q0itfinance = data.get('q0itfinance') or " "
    q0rpo = data.get('q0rpo') or " "
    q0poa = data.get('q0poa') or " "
    q0mspi = data.get('q0mspi') or " "
    q0msl = data.get('q0msl') or " "
    q0fqmr = data.get('q0fqmr') or " "
    q0dcap = data.get('q0dcap') or " "
    q0cifw = data.get('q0cifw') or " "
    valsubmissionid = data.get('valsubmissionid') or " "
    
   
    

    totalcostval = (int(str(donutit).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutrpo).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutpoa).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutdcap).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutcip).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutmspi).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutmsl).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutfqmr).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutcifw).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutoem).replace('£','').replace(',','').replace(' ','')))
    per1x=0
    per2x=0
    per3x=0
    per4x=0
    per5x=0
    per6x=0
    per7x=0
    per8x=0
    per9x=0
    per10x=0 

    
    if totalcostval!=0:
        per1x = round((int(str(donutit).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per2x = round((int(str(donutrpo).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per3x = round((int(str(donutpoa).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per4x = round((int(str(donutdcap).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per5x = round((int(str(donutcip).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per6x = round((int(str(donutmspi).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per7x = round((int(str(donutmsl).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per8x = round((int(str(donutfqmr).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per9x = round((int(str(donutcifw).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per10x = round((int(str(donutoem).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
    
    if not client_name:
        return "Error: 'client_name' parameter is required", 400

    zip_dosya = 'template.zip'
    gecici_zip_dosya = 'temp_template.zip'

    output_ppt_path = 'output.pptx'

    # Dosya mevcutsa sil
    if os.path.exists(output_ppt_path):
        os.remove(output_ppt_path)


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
                                donutcifw=donutcifw,donutoem=donutoem,
                                prpoval=prpoval,ppoaval=ppoaval,pcipval=pcipval,pmspival=pmspival,pmslval=pmslval,pfqmrval=pfqmrval,
                                pdcapval=pdcapval,pcifwval=pcifwval,poemval=poemval,pitfinanceval=pitfinanceval,totalcostval=totalcostval,per1x=per1x,
                                per2x=per2x,per3x=per3x,per4x=per4x,per5x=per5x,per6x=per6x,per7x=per7x,per8x=per8x,per9x=per9x,per10x=per10x,
                                asrpoval=asrpoval,aspoaval=aspoaval,ascipval=ascipval,asmspival=asmspival,asmslval=asmslval,asfqmrval=asfqmrval,
                                asdcapval=asdcapval,ascifwval=ascifwval,asoemval=asoemval,asitfinance=asitfinance,eytimerpo=eytimerpo,
                                eytimepoa=eytimepoa,eytimecip=eytimecip,eytimemspi=eytimemspi,eytimemsl=eytimemsl,eytimefqmr=eytimefqmr,
                                eytimedcap=eytimedcap,eytimecifw=eytimecifw,eytimeoem=eytimeoem,eytimeitfinance=eytimeitfinance,
                                q1itfinance=q1itfinance,q1rpo=q1rpo,q2rpo=q2rpo,q3rpo=q3rpo,q1poa=q1poa,q2poa=q2poa,q3poa=q3poa,q1cip=q1cip,
    q2cip=q2cip,q3cip=q3cip,q4cip=q4cip,q1mspi=q1mspi,q2mspi=q2mspi,q3mspi=q3mspi,q1msl=q1msl,q2msl=q2msl,q1fqmr=q1fqmr,
    q2fqmr=q2fqmr,q3fqmr=q3fqmr,q4fqmr=q4fqmr,q1dcap=q1dcap,q2dcap=q2dcap,q3dcap=q3dcap,q1cifw=q1cifw,q2cifw=q2cifw,q0itfinance=q0itfinance,q0rpo=q0rpo,q0poa=q0poa,q0mspi=q0mspi,q0msl=q0msl,q0fqmr=q0fqmr,q0dcap=q0dcap,q0cifw=q0cifw,
    valsubmissionid=valsubmissionid)
    
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

@app.route('/create-client-ppt', methods=['POST'])
def create_client_ppt():
    # 'client_name' parametresini POST isteği ile al
    data = request.get_json()
    firstname = data.get('firstname') or ""
    lastname = data.get('lastname') or ""
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

    
    

    prpoval = data.get('prpoval') or "0"
    ppoaval = data.get('ppoaval') or "0"
    pcipval = data.get('pcipval') or "0"
    pmspival = data.get('pmspival') or "0"
    pmslval = data.get('pmslval') or "0"
    pfqmrval = data.get('pfqmrval') or "0"
    pdcapval = data.get('pdcapval') or "0"
    pcifwval = data.get('pcifwval') or "0"
    poemval = data.get('poemval') or "0"
    pitfinanceval = data.get('pitfinanceval') or "0"


    valsubmissionid = data.get('valsubmissionid') or " "
    q0mspi = data.get('q0mspi') or " "
    q1mspi = data.get('q1mspi') or " "
    q2mspi = data.get('q2mspi') or " "
    q3mspi = data.get('q3mspi') or " "
    q0fqmr = data.get('q0fqmr') or " "
    q1fqmr = data.get('q1fqmr') or " "
    q2fqmr = data.get('q2fqmr') or " "
    q3fqmr = data.get('q3fqmr') or " "
    q4fqmr = data.get('q4fqmr') or " "
    q0dcap = data.get('q0dcap') or " "
    q1dcap = data.get('q1dcap') or " "
    q2dcap = data.get('q2dcap') or " "
    q3dcap = data.get('q3dcap') or " "
    

    
    
    totalcostval = (int(str(donutit).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutrpo).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutpoa).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutdcap).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutcip).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutmspi).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutmsl).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutfqmr).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutcifw).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutoem).replace('£','').replace(',','').replace(' ','')))
    per1x=0
    per2x=0
    per3x=0
    per4x=0
    per5x=0
    per6x=0
    per7x=0
    per8x=0
    per9x=0
    per10x=0 

    client_name = firstname + " " + lastname

    
    if totalcostval!=0:
        per1x = round((int(str(donutit).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per2x = round((int(str(donutrpo).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per3x = round((int(str(donutpoa).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per4x = round((int(str(donutcifw).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per5x = round((int(str(donutcip).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per6x = round((int(str(donutmspi).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per7x = round((int(str(donutmsl).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per8x = round((int(str(donutfqmr).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per9x = round((int(str(donutdcap).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per10x = round((int(str(donutoem).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
    

    zip_dosya = 'client_template.zip'
    gecici_zip_dosya = 'temp_client_template.zip'

    


    zip_path = r"client_template.zip"  # Tam dosya yolunu girin
    output_pptx_path = r"client_output.pptx"  # Çıkış dosyasının yolunu belirtin

    modify_slide_xml_and_image_client(zip_path, output_pptx_path,client_name,itfinance,rpo,poa,cip,mspi,valmsl,valfqmr,valdcap,
                               valcifw,valoem,valbnft,valnpvv,valacd,valroi,valinvestment,valmonths,valhours,
                               year1invest=year1invest,
                                year1return=year1total,year2invest=year2invest,year2return=year2otal,year3invest=year3invest,
                                year3return=year3total,year4invest=year4invest,year4return=year4total,year5invest=year5invest,
                                year5return=year5total,costofdoingnothing1=costofdoingnothing1,itfinanceper=itfinanceper,rpoper=rpoper,
                                poaper=poaper,cipper=cipper,mspiper=mspiper,valmslper=valmslper,valfqmrper=valfqmrper,valdcapper=valdcapper,
                                valcifwper=valcifwper,valoemper=valoemper,donutit=donutit,donutrpo=donutrpo,donutpoa=donutpoa,
                                donutdcap=donutdcap,donutcip=donutcip,donutmspi=donutmspi,donutmsl=donutmsl,donutfqmr=donutfqmr,
                                donutcifw=donutcifw,donutoem=donutoem,
                                prpoval=prpoval,ppoaval=ppoaval,pcipval=pcipval,pmspival=pmspival,pmslval=pmslval,pfqmrval=pfqmrval,
                                pdcapval=pdcapval,pcifwval=pcifwval,poemval=poemval,pitfinanceval=pitfinanceval,totalcostval=totalcostval,
                                per1x=per1x,per2x=per2x,per3x=per3x,per4x=per4x,per5x=per5x,per6x=per6x,per7x=per7x,per8x=per8x,per9x=per9x,per10x=per10x,
                                valsubmissionid=valsubmissionid,q0mspi=q0mspi,q1mspi=q1mspi,q2mspi=q2mspi,q3mspi=q3mspi,q0fqmr=q0fqmr,q1fqmr=q1fqmr,
    q2fqmr=q2fqmr,q3fqmr=q3fqmr,q4fqmr=q4fqmr,q0dcap=q0dcap,q1dcap=q1dcap,q2dcap=q2dcap,q3dcap=q3dcap)
    
    zip_path = 'client_template.zip'  # Güncellemek istediğin template.zip
    output_zip_path = 'client_template.zip'  # Çıkış dosyasının adı

    

   


    pptx_io = io.BytesIO()
    with open(output_pptx_path, 'rb') as f:
        pptx_io.write(f.read())
    pptx_io.seek(0)

    # Dosyayı base64 formatında encode et
    pptx_base64 = base64.b64encode(pptx_io.read()).decode('utf-8')

    # JSON formatında base64 ile encode edilmiş dosyayı döndür
    return jsonify({
        'file_name': 'client_presentation.pptx',
        'file_content': pptx_base64
    })


@app.route('/create-questionare-ppt', methods=['POST'])
def create_questionare_ppt():
    # 'client_name' parametresini POST isteği ile al
    data = request.get_json()
    client_name = data.get('client_name') or ""
    baitval = data.get('baitval') or "0"
    barpoval = data.get('barpoval') or "0"
    bapoaval = data.get('bapoaval') or "0"
    bacipval = data.get('bacipval') or "0"
    bamspival = data.get('bamspival') or "0"
    bamslval = data.get('bamslval') or "0"
    bafqmrval = data.get('bafqmrval') or "0"
    badcapval = data.get('badcapval') or "0"
    bacifwval = data.get('bacifwval') or "0"
    baoemval = data.get('baoemval') or "0"
    batotalval = data.get('batotalval') or "0"

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

    period1 = data.get('period1') or " "
    period2 = data.get('period2') or " "

    
    

    prpoval = data.get('prpoval') or "0"
    ppoaval = data.get('ppoaval') or "0"
    pcipval = data.get('pcipval') or "0"
    pmspival = data.get('pmspival') or "0"
    pmslval = data.get('pmslval') or "0"
    pfqmrval = data.get('pfqmrval') or "0"
    pdcapval = data.get('pdcapval') or "0"
    pcifwval = data.get('pcifwval') or "0"
    poemval = data.get('poemval') or "0"
    pitfinanceval = data.get('pitfinanceval') or "0"
    
    totalcostval = (int(str(donutit).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutrpo).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutpoa).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutdcap).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutcip).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutmspi).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutmsl).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutfqmr).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutcifw).replace('£','').replace(',','').replace(' ',''))+
                    int(str(donutoem).replace('£','').replace(',','').replace(' ','')))
    per1x=0
    per2x=0
    per3x=0
    per4x=0
    per5x=0
    per6x=0
    per7x=0
    per8x=0
    per9x=0
    per10x=0 


    
    if totalcostval!=0:
        per1x = round((int(str(donutit).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per2x = round((int(str(donutrpo).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per3x = round((int(str(donutpoa).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per4x = round((int(str(donutdcap).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per5x = round((int(str(donutcip).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per6x = round((int(str(donutmspi).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per7x = round((int(str(donutmsl).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per8x = round((int(str(donutfqmr).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per9x = round((int(str(donutcifw).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
        per10x = round((int(str(donutoem).replace('£','').replace(',','').replace(' ',''))/totalcostval*100),1)
    

    zip_dosya = 'questionare_template.zip'
    gecici_zip_dosya = 'temp_questionare_template.zip'

    output_ppt_path = 'questionare_output.pptx'

    # Dosya mevcutsa sil
    if os.path.exists(output_ppt_path):
        os.remove(output_ppt_path)


    zip_path = r"questionare_template.zip"  # Tam dosya yolunu girin
    output_pptx_path = r"questionare_output.pptx"  # Çıkış dosyasının yolunu belirtin

    modify_slide_xml_and_image_questionare(zip_path, output_pptx_path,client_name,
                                           baitval=baitval,barpoval=barpoval,bapoaval=bapoaval,bacipval=bacipval,bamspival=bamspival,bamslval=bamslval,bafqmrval=bafqmrval,badcapval=badcapval,
                               bacifwval=bacifwval,baoemval=baoemval,batotalval=batotalval,
                               donutit=donutit,donutrpo=donutrpo,donutpoa=donutpoa,donutdcap=donutdcap,donutcip=donutcip,donutmspi=donutmspi,donutmsl=donutmsl,
                               donutfqmr=donutfqmr,donutcifw=donutcifw,donutoem=donutoem,prpoval=prpoval,ppoaval=ppoaval,pcipval=pcipval,pmspival=pmspival,pmslval=pmslval,
                               pfqmrval=pfqmrval,pdcapval=pdcapval,pcifwval=pcifwval,poemval=poemval,pitfinanceval=pitfinanceval,period1=period1,period2=period2)
    
    zip_path = 'questionare_template.zip'  # Güncellemek istediğin template.zip
    output_zip_path = 'questionare_template.zip'  # Çıkış dosyasının adı

    

   


    pptx_io = io.BytesIO()
    with open(output_pptx_path, 'rb') as f:
        pptx_io.write(f.read())
    pptx_io.seek(0)

    # Dosyayı base64 formatında encode et
    pptx_base64 = base64.b64encode(pptx_io.read()).decode('utf-8')

    # JSON formatında base64 ile encode edilmiş dosyayı döndür
    return jsonify({
        'file_name': 'questionare_presentation.pptx',
        'file_content': pptx_base64
    })


@app.route('/get-emails', methods=['POST'])
def getemails():
    
    # API anahtarınızı buraya ekleyin
    api_key = "cc4c65854e01a444f918600ff6d3e4be"
    data = request.get_json()
    firstemail = data.get('firstemail')
    tableid = data.get('tableid')
    # GET isteği gönderme
    url = "https://eu-api.jotform.com/form/{}/submissions".format(tableid)
    response = requests.get(url, params={'apiKey': api_key})

    datares = response.json()
    emails=firstemail
    for i in datares['content']:
        if i['answers']['4']:
            emails+=';'
            emails+=i['answers']['4']['answer']
            

    return emails

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=8080)
