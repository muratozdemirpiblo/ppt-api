import zipfile
import os
import xml.etree.ElementTree as ET
import requests
import shutil

def modify_slide_xml_and_image(zip_path, output_pptx_path):
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
                elem.text = elem.text.replace('valclient', 'test')
            if 'itfinance' in elem.text:
                elem.text = elem.text.replace('itfinance', '34,340')
            if 'rpo' in elem.text:
                elem.text = elem.text.replace('rpo', '34,340')
            if 'poa' in elem.text:
                elem.text = elem.text.replace('poa', '34,340')
            if 'cip' in elem.text:
                elem.text = elem.text.replace('cip', '34,340')
            if 'mspi' in elem.text:
                elem.text = elem.text.replace('mspi', '34,340')
            if 'valmsl' in elem.text:
                elem.text = elem.text.replace('valmsl', '34,340')
            if 'valfqmr' in elem.text:
                elem.text = elem.text.replace('valfqmr', '34,340')
            if 'valdcap' in elem.text:
                elem.text = elem.text.replace('valdcap', '34,340')
            if 'valcifw' in elem.text:
                elem.text = elem.text.replace('valcifw', '34,340')
            if 'valoem' in elem.text:
                elem.text = elem.text.replace('valoem', '34,340')
            if 'valbnft' in elem.text:
                elem.text = elem.text.replace('valbnft', '£34,340')
            if '£valnpvv' in elem.text:
                elem.text = elem.text.replace('£valnpvv', '£34,340')
            if 'valacd' in elem.text:
                elem.text = elem.text.replace('valacd', '34,340')
            if 'valroi' in elem.text:
                elem.text = elem.text.replace('valroi', '34')
            if 'valinvestment' in elem.text:
                elem.text = elem.text.replace('valinvestment', '34,340')
            if 'valmonths' in elem.text:
                elem.text = elem.text.replace('valmonths', '2')

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

# Kullanım
zip_path = r"template.zip"  # Tam dosya yolunu girin
output_pptx_path = r"output.pptx"  # Çıkış dosyasının yolunu belirtin

modify_slide_xml_and_image(zip_path, output_pptx_path)
