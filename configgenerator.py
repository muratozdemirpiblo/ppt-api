import zipfile
import os
import xml.etree.ElementTree as ET
import requests
import shutil

def modify_slide_xml_and_image(zip_path, output_pptx_path, values, image_url):
    # Geçici çalışma dizinini oluştur
    temp_dir = 'temp_pptx'
    os.makedirs(temp_dir, exist_ok=True)

    # .pptx dosyasını aç ve dosyaları çıkar
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    # slides klasöründeki tüm slide XML dosyalarını bul ve işle
    slides_dir = os.path.join(temp_dir, 'ppt', 'slides')
    slide_files = [f for f in os.listdir(slides_dir) if f.startswith('slide') and f.endswith('.xml')]
    
    # Değerler listesi için index
    value_index = 0

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
            if '£XX,000' in elem.text:
                elem.text = elem.text.replace('£XX,000', values[value_index])
                value_index += 1
                if value_index >= len(values):
                    break  # Listedeki tüm değerler kullanıldığında durdur

        # Güncellenmiş slide XML dosyasını kaydet
        tree.write(slide_xml_path, xml_declaration=True, encoding='UTF-8')

        if value_index >= len(values):
            break  # Tüm değerler değiştirildiyse döngüyü durdur

    # Görseli güncelleme
    image_path = os.path.join(temp_dir, 'ppt', 'media', 'image16.png')

    # URL'den yeni görseli indir
    response = requests.get(image_url, stream=True)
    if response.status_code == 200:
        # İndirilen görseli ppt/media klasöründeki image16.png ile değiştir
        with open(image_path, 'wb') as out_file:
            shutil.copyfileobj(response.raw, out_file)
        print("image16.png başarıyla indirildi ve değiştirildi.")
    else:
        print("Görsel URL'den indirilemedi.")

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
values = ['£220,155', '£315,400', '£125,600', '£400,250', '£540,000', '£155,300', '£230,120', '£180,450', '£260,700', '£310,500']
zip_path = r"template.zip"  # Tam dosya yolunu girin
output_pptx_path = r"output.pptx"  # Çıkış dosyasının yolunu belirtin
image_url = 'https://muratozdemirpiblo.github.io/Background.jpg'  # İlgili görselin URL'sini buraya girin

modify_slide_xml_and_image(zip_path, output_pptx_path, values, image_url)
