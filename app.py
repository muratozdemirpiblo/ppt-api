from flask import Flask, jsonify
from pptx import Presentation
from pptx.util import Inches
import os
import io
import base64

app = Flask(__name__)

@app.route('/create-ppt', methods=['POST'])
def create_ppt():
    # PowerPoint sunumu oluştur
    client_name = request.form.get('client_name')


    prs = Presentation()

    # İlk slaytı ekle
    slide_layout = prs.slide_layouts[5]  # Boş slayt
    slide = prs.slides.add_slide(slide_layout)

    # Arkaplan resmi ekleme
    img_path_bg = os.path.join(os.getcwd(), 'slideassets', 'background1.png')
    slide.shapes.add_picture(img_path_bg, Inches(0), Inches(0), width=Inches(13.3), height=Inches(7.5))

    # Diğer resmi ekleme
    img_path = os.path.join(os.getcwd(), 'slideassets', 'image1.jpg')
    slide.shapes.add_picture(img_path, Inches(7.61), Inches(1.40), width=Inches(4.66), height=Inches(4.67))

    # Başlık ekleme
    left = Inches(0.75)
    top = Inches(1.67)
    width = Inches(12)
    height = Inches(1.5)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    p = text_frame.add_paragraph()
    p.text = "Client Name:\nValue Board Pack template"
    p.font.size = Inches(0.64)  # Yaklaşık 48 pt
    p.font.bold = True

    # Başlık ekleme
    left = Inches(0,75)
    top = Inches(4,36)
    width = Inches(5,78)
    height = Inches(0,68)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    p = text_frame.add_paragraph()
    p.text = "Financials"
    p.font.size = Inches(0.28)  # Yaklaşık 48 pt
   

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

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=8080)
