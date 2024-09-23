import zipfile
import os

def update_xml_text_in_zip(zip_path, xml_path):
    # Geçici dosya için bir dizin oluştur
    temp_dir = 'temp_dir'
    os.makedirs(temp_dir, exist_ok=True)

    # ZIP dosyasını aç ve içeriğini çıkar
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    # XML dosyasını metin olarak oku
    xml_file_path = os.path.join(temp_dir, xml_path)
    with open(xml_file_path, 'r', encoding='utf-8') as xml_file:
        xml_content = xml_file.read()

    # Metin olarak düzenleme (örneğin <c:numCache> ekleme)
    new_content = xml_content.replace('<c:numCache></c:numCache>', '''
<c:numCache>
    <c:formatCode>_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)</c:formatCode>
    <c:ptCount val="14"/>
    <c:pt idx="0"><c:v>33000</c:v></c:pt>
    <c:pt idx="1"><c:v>66430.547826086971</c:v></c:pt>
    <c:pt idx="2"><c:v>29400</c:v></c:pt>
    <c:pt idx="3"><c:v>80665.665217391332</c:v></c:pt>
    <c:pt idx="4"><c:v>30870</c:v></c:pt>
    <c:pt idx="5"><c:v>94900.782608695677</c:v></c:pt>
    <c:pt idx="6"><c:v>32413.5</c:v></c:pt>
    <c:pt idx="7"><c:v>94900.782608695677</c:v></c:pt>
    <c:pt idx="8"><c:v>34034.175000000003</c:v></c:pt>
    <c:pt idx="9"><c:v>94900.782608695677</c:v></c:pt>
</c:numCache>
''')

    # Güncellenmiş içeriği tekrar dosyaya yaz
    with open(xml_file_path, 'w', encoding='utf-8') as xml_file:
        xml_file.write(new_content)

    # Güncellenmiş dosyaları tekrar ZIP yap
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
        for foldername, subfolders, filenames in os.walk(temp_dir):
            for filename in filenames:
                filepath = os.path.join(foldername, filename)
                arcname = os.path.relpath(filepath, temp_dir)
                zip_ref.write(filepath, arcname)

    # Geçici dizini temizle
    for foldername, subfolders, filenames in os.walk(temp_dir, topdown=False):
        for filename in filenames:
            os.remove(os.path.join(foldername, filename))
        os.rmdir(foldername)

# Kullanım
zip_path = 'template.zip'
xml_path = 'ppt/charts/chart1.xml'  # Güncellenecek XML dosyasının yolu
update_xml_text_in_zip(zip_path, xml_path)
