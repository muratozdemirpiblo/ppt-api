from flask import Flask, request, send_file, jsonify
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import os
import io
import base64

app = Flask(__name__)

@app.route('/create-ppt', methods=['POST'])
def create_ppt():
    # 'client_name' parametresini POST isteği ile al
    data = request.get_json()
    client_name = data.get('client_name')
    
    if not client_name:
        return "Error: 'client_name' parameter is required", 400

    # PowerPoint sunumu oluştur
    prs = Presentation()

    # Sunum düzenini geniş ekran (16:9) olarak ayarla
    prs.slide_width = Inches(13.33)  # 16:9 formatı için genişlik
    prs.slide_height = Inches(7.5)   # 16:9 formatı için yükseklik

    # İlk slaytı ekle
    slide_layout = prs.slide_layouts[5]  # Boş slayt
    slide = prs.slides.add_slide(slide_layout)

    # Arkaplan resmi ekleme
    img_path_bg = os.path.join(os.getcwd(), 'slideassets', 'background1.png')
    slide.shapes.add_picture(img_path_bg, Inches(0), Inches(0), width=Inches(13.33), height=Inches(7.5))

    # Diğer resmi ekleme
    img_path = os.path.join(os.getcwd(), 'slideassets', 'image1.jpg')
    slide.shapes.add_picture(img_path, Inches(7.61), Inches(1.40), width=Inches(4.66), height=Inches(4.67))

    # Başlık ekleme - Client Name'den sonra yeni satır eklemek için \n kullanıyoruz
    left = Inches(0.75)
    top = Inches(1.67)
    width = Inches(12)
    height = Inches(1.5)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    p = text_frame.add_paragraph()
    p.text = f"Client Name:\n{client_name}"  # Client Name'den sonra yeni satır
    p.font.size = Inches(0.64)  # Yaklaşık 48 pt
    p.font.bold = True

    # Başlık ekleme
    left = Inches(0.75)
    top = Inches(4.36)
    width = Inches(5.78)
    height = Inches(0.68)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    p = text_frame.add_paragraph()
    p.text = "Financials"
    p.font.size = Inches(0.28)  # Yaklaşık 48 pt

    # İkinci slaytı ekle
    slide_layout2 = prs.slide_layouts[5]  # Başlık ve içerik düzeni
    slide2 = prs.slides.add_slide(slide_layout2)

    for shape in slide2.shapes:
        if not shape.has_text_frame:
            continue
        text_frame = shape.text_frame
        if text_frame.text == 'Click to add title' or text_frame.text == 'Click to add text':
            sp = shape
            slide2.shapes._spTree.remove(sp._element)

    # Üst kutucuk (Başlık) ekleme
    left = Inches(0.75)
    top = Inches(0.75)
    width = Inches(12)
    height = Inches(1.0)
    textbox = slide2.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    p = text_frame.add_paragraph()
    p.text = "Calculator headings categorised into the below sections"
    p.font.size = Pt(24)  # Yaklaşık 24 pt
    p.font.name = 'Montserrat SemiBold'
    p.font.color.rgb = RGBColor(0, 0, 0)  # Siyah renk

    # Dosyayı belleğe kaydet
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)

    # Dosyayı base64 formatında encode et
    pptx_base64 = base64.b64encode(pptx_io.read()).decode('utf-8')

    # JSON formatında base64 ile encode edilmiş dosyayı döndür
    return jsonify({
        'file_name': 'presentation.pptx',
        'file_content': pptx_base64
    })

@app.route('/test', methods=['GET'])
def create_ppt_test():
    # PowerPoint sunumu oluştur
    prs = Presentation()

    # Sunum düzenini geniş ekran (16:9) olarak ayarla
    prs.slide_width = Inches(13.33)  # 16:9 formatı için genişlik
    prs.slide_height = Inches(7.5)   # 16:9 formatı için yükseklik

    # İlk slaytı ekle
    slide_layout = prs.slide_layouts[5]  # Boş slayt
    slide = prs.slides.add_slide(slide_layout)

    # Arkaplan resmi ekleme
    img_path_bg = os.path.join(os.getcwd(), 'slideassets', 'background1.png')
    slide.shapes.add_picture(img_path_bg, Inches(0), Inches(0), width=Inches(13.33), height=Inches(7.5))

    # Diğer resmi ekleme
    img_path = os.path.join(os.getcwd(), 'slideassets', 'image1.jpg')
    slide.shapes.add_picture(img_path, Inches(7.61), Inches(1.40), width=Inches(4.66), height=Inches(4.67))

    # Başlık ekleme - Client Name'den sonra yeni satır eklemek için \n kullanıyoruz
    left = Inches(0.75)
    top = Inches(1.67)
    width = Inches(12)
    height = Inches(1.5)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    p = text_frame.add_paragraph()
    p.font.name = 'Montserrat SemiBold'
    p.text = "Client Name:\nValue Board Pack template"
    p.font.size = Inches(0.64)  # Yaklaşık 48 pt
    p.font.bold = True

    # Başlık ekleme
    left = Inches(0.75)
    top = Inches(4.36)
    width = Inches(5.78)
    height = Inches(0.68)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    p.font.name = 'Montserrat SemiBold'
    p = text_frame.add_paragraph()
    p.text = "Financials"
    p.font.size = Inches(0.28)  # Yaklaşık 48 pt



    # İkinci slaytı ekle
    slide_layout2 = prs.slide_layouts[5]  # Başlık ve içerik düzeni
    slide2 = prs.slides.add_slide(slide_layout2)

    for shape in slide2.shapes:
        if not shape.has_text_frame:
            continue
        text_frame = shape.text_frame
        if text_frame.text == 'Click to add title' or text_frame.text == 'Click to add text':
            sp = shape
            slide2.shapes._spTree.remove(sp._element)

    # Üst kutucuk (Başlık) ekleme
    left = Inches(0.75)
    top = Inches(0.75)
    width = Inches(12)
    height = Inches(1.0)
    textbox = slide2.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    p = text_frame.add_paragraph()
    p.text = "Calculator headings categorised into the below sections"
    p.font.size = Pt(24)  # Yaklaşık 24 pt
    p.font.name = 'Montserrat SemiBold'
    p.font.color.rgb = RGBColor(0, 0, 0)  # Siyah renk


    # Dosyayı belleğe kaydet
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)

    # Dosyayı doğrudan yanıt olarak döndür
    return send_file(pptx_io, download_name='test_presentation.pptx', as_attachment=True)

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=8080)
