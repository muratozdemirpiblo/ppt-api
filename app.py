from flask import Flask, send_file
from pptx import Presentation
from pptx.util import Inches
import os

app = Flask(__name__)

@app.route('/create-ppt', methods=['GET'])
def create_ppt():
    # PowerPoint sunumu oluştur
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
    p.text = "Client Name: Value Board Pack template"
    p.font.size = Inches(0.64)  # Yaklaşık 48 pt
    p.font.bold = True

    # Sunumu kaydet
    pptx_file = 'example.pptx'
    prs.save(pptx_file)

    # Sunumu indirme olarak döndür
    return send_file(pptx_file, as_attachment=True)

if __name__ == '__main__':
    # slideassets klasörünün mevcut olduğundan emin olun
    if not os.path.exists('slideassets'):
        os.makedirs('slideassets')
    # Flask sunucusunu başlat
    app.run(debug=True)


    
