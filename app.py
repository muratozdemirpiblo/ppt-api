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
    

@app.route('/create-ppt', methods=['POST'])
def create_ppt():
    # 'client_name' parametresini POST isteği ile al
    data = request.get_json()
    client_name = data.get('client_name')
    
    if not client_name:
        return "Error: 'client_name' parameter is required", 400


    # Dosyayı belleğe kaydet
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)

    values = ['£220,155', '£315,400', '£125,600', '£400,250', '£540,000', '£155,300', '£230,120', '£180,450', '£260,700', '£310,500']
    zip_path = r"template.zip"  # Tam dosya yolunu girin
    output_pptx_path = r"output.pptx"  # Çıkış dosyasının yolunu belirtin
    image_url = 'https://muratozdemirpiblo.github.io/Background.jpg'  # İlgili görselin URL'sini buraya girin

    modify_slide_xml_and_image(zip_path, output_pptx_path, values, image_url)

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


@app.route('/chart')
def donut_chart():
    return render_template('sales.html')

@app.route('/chart-image')
def get_chart_image():
    # Chrome seçeneklerini ayarla
    options = webdriver.ChromeOptions()
    options.add_argument('headless')  # Arayüzsüz modda çalıştır
    options.add_argument('window-size=480x480')  # Pencere boyutu

    # Chrome sürücüsünü başlat
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

    # Flask uygulamasındaki / sayfasına git
    driver.get("http://127.0.0.1:8080/chart")

    # Grafik yüklenmesi için bekle
    time.sleep(2)

    # Ekran görüntüsü al
    screenshot_path = "chart_screenshot.png"
    driver.save_screenshot(screenshot_path)

    # Tarayıcıyı kapat
    driver.quit()

    # Ekran görüntüsünü Pillow ile işleyip grafiği kırp
    image = Image.open(screenshot_path)

    # Grafik alanını doğru ayarlamak için kırpma koordinatlarını dinamik olarak belirleyin
    # Örnek olarak, görüntü boyutunu kontrol edelim
    width, height = image.size

    # Grafik alanının boyutunu ve konumunu belirlemek için aşağıdaki değerleri ayarlayın
    # Bu değerleri, sayfanın HTML ve CSS yapısına göre güncellemeniz gerekebilir
    chart_left = int(width * 0.1)  # Sol kenar boşluğu
    chart_top = int(height * 0.1)  # Üst kenar boşluğu
    chart_right = int(width * 0.9)  # Sağ kenar boşluğu
    chart_bottom = int(height * 0.9)  # Alt kenar boşluğu

    # Kırpma işlemi
    cropped_image = image.crop((chart_left, chart_top, chart_right, chart_bottom))

    # Kırpılmış resmi kaydet
    cropped_image.save(screenshot_path)

    # Resmi döndür
    return send_file(screenshot_path, mimetype='image/png')




if __name__ == "__main__":
    app.run(host='0.0.0.0', port=8080)
