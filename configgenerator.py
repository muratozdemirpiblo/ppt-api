import zipfile
import os
import shutil

def create_chart_xml():
    # XML içeriğini tanımla
    xml_content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace
	xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
	xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
	xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart">
	<c:date1904 val="0"/>
	<c:lang val="tr-TR"/>
	<c:roundedCorners val="0"/>
	<mc:AlternateContent
		xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
		<mc:Choice Requires="c14"
			xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">
			<c14:style val="102"/>
		</mc:Choice>
		<mc:Fallback>
			<c:style val="2"/>
		</mc:Fallback>
	</mc:AlternateContent>
	<c:chart>
		<c:autoTitleDeleted val="1"/>
		<c:plotArea>
			<c:layout>
				<c:manualLayout>
					<c:layoutTarget val="inner"/>
					<c:xMode val="edge"/>
					<c:yMode val="edge"/>
					<c:x val="5.442140890092332E-2"/>
					<c:y val="3.108893616642484E-2"/>
					<c:w val="0.92859981406902237"/>
					<c:h val="0.85615280534809812"/>
				</c:manualLayout>
			</c:layout>
			<c:barChart>
				<c:barDir val="col"/>
				<c:grouping val="clustered"/>
				<c:varyColors val="0"/>
				<c:ser>
					<c:idx val="0"/>
					<c:order val="0"/>
					<c:spPr>
						<a:gradFill rotWithShape="1">
							<a:gsLst>
								<a:gs pos="0">
									<a:schemeClr val="accent1">
										<a:satMod val="103000"/>
										<a:lumMod val="102000"/>
										<a:tint val="94000"/>
									</a:schemeClr>
								</a:gs>
								<a:gs pos="50000">
									<a:schemeClr val="accent1">
										<a:satMod val="110000"/>
										<a:lumMod val="100000"/>
										<a:shade val="100000"/>
									</a:schemeClr>
								</a:gs>
								<a:gs pos="100000">
									<a:schemeClr val="accent1">
										<a:lumMod val="99000"/>
										<a:satMod val="120000"/>
										<a:shade val="78000"/>
									</a:schemeClr>
								</a:gs>
							</a:gsLst>
							<a:lin ang="5400000" scaled="0"/>
						</a:gradFill>
						<a:ln>
							<a:noFill/>
						</a:ln>
						<a:effectLst>
							<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
								<a:srgbClr val="000000">
									<a:alpha val="63000"/>
								</a:srgbClr>
							</a:outerShdw>
						</a:effectLst>
					</c:spPr>
					<c:invertIfNegative val="0"/>
					<c:dPt>
						<c:idx val="0"/>
						<c:invertIfNegative val="0"/>
						<c:bubble3D val="0"/>
						<c:spPr>
							<a:solidFill>
								<a:schemeClr val="accent1"/>
							</a:solidFill>
							<a:ln>
								<a:noFill/>
							</a:ln>
							<a:effectLst>
								<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
									<a:srgbClr val="000000">
										<a:alpha val="63000"/>
									</a:srgbClr>
								</a:outerShdw>
							</a:effectLst>
						</c:spPr>
						<c:extLst>
							<c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"
								xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">
								<c16:uniqueId val="{00000001-83F9-4BC0-8CB9-9896CA7D3BE8}"/>
							</c:ext>
						</c:extLst>
					</c:dPt>
					<c:dPt>
						<c:idx val="1"/>
						<c:invertIfNegative val="0"/>
						<c:bubble3D val="0"/>
						<c:spPr>
							<a:solidFill>
								<a:schemeClr val="accent3"/>
							</a:solidFill>
							<a:ln>
								<a:noFill/>
							</a:ln>
							<a:effectLst>
								<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
									<a:srgbClr val="000000">
										<a:alpha val="63000"/>
									</a:srgbClr>
								</a:outerShdw>
							</a:effectLst>
						</c:spPr>
						<c:extLst>
							<c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"
								xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">
								<c16:uniqueId val="{00000003-83F9-4BC0-8CB9-9896CA7D3BE8}"/>
							</c:ext>
						</c:extLst>
					</c:dPt>
					<c:dPt>
						<c:idx val="3"/>
						<c:invertIfNegative val="0"/>
						<c:bubble3D val="0"/>
						<c:spPr>
							<a:solidFill>
								<a:schemeClr val="accent2"/>
							</a:solidFill>
							<a:ln>
								<a:noFill/>
							</a:ln>
							<a:effectLst>
								<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
									<a:srgbClr val="000000">
										<a:alpha val="63000"/>
									</a:srgbClr>
								</a:outerShdw>
							</a:effectLst>
						</c:spPr>
						<c:extLst>
							<c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"
								xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">
								<c16:uniqueId val="{00000005-83F9-4BC0-8CB9-9896CA7D3BE8}"/>
							</c:ext>
						</c:extLst>
					</c:dPt>
					<c:dPt>
						<c:idx val="4"/>
						<c:invertIfNegative val="0"/>
						<c:bubble3D val="0"/>
						<c:spPr>
							<a:solidFill>
								<a:schemeClr val="accent3"/>
							</a:solidFill>
							<a:ln>
								<a:noFill/>
							</a:ln>
							<a:effectLst>
								<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
									<a:srgbClr val="000000">
										<a:alpha val="63000"/>
									</a:srgbClr>
								</a:outerShdw>
							</a:effectLst>
						</c:spPr>
						<c:extLst>
							<c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"
								xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">
								<c16:uniqueId val="{00000007-83F9-4BC0-8CB9-9896CA7D3BE8}"/>
							</c:ext>
						</c:extLst>
					</c:dPt>
					<c:dPt>
						<c:idx val="6"/>
						<c:invertIfNegative val="0"/>
						<c:bubble3D val="0"/>
						<c:spPr>
							<a:solidFill>
								<a:schemeClr val="accent2"/>
							</a:solidFill>
							<a:ln>
								<a:noFill/>
							</a:ln>
							<a:effectLst>
								<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
									<a:srgbClr val="000000">
										<a:alpha val="63000"/>
									</a:srgbClr>
								</a:outerShdw>
							</a:effectLst>
						</c:spPr>
						<c:extLst>
							<c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"
								xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">
								<c16:uniqueId val="{00000009-83F9-4BC0-8CB9-9896CA7D3BE8}"/>
							</c:ext>
						</c:extLst>
					</c:dPt>
					<c:dPt>
						<c:idx val="7"/>
						<c:invertIfNegative val="0"/>
						<c:bubble3D val="0"/>
						<c:spPr>
							<a:solidFill>
								<a:schemeClr val="accent3"/>
							</a:solidFill>
							<a:ln>
								<a:noFill/>
							</a:ln>
							<a:effectLst>
								<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
									<a:srgbClr val="000000">
										<a:alpha val="63000"/>
									</a:srgbClr>
								</a:outerShdw>
							</a:effectLst>
						</c:spPr>
						<c:extLst>
							<c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"
								xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">
								<c16:uniqueId val="{0000000B-83F9-4BC0-8CB9-9896CA7D3BE8}"/>
							</c:ext>
						</c:extLst>
					</c:dPt>
					<c:dPt>
						<c:idx val="9"/>
						<c:invertIfNegative val="0"/>
						<c:bubble3D val="0"/>
						<c:spPr>
							<a:solidFill>
								<a:schemeClr val="accent2"/>
							</a:solidFill>
							<a:ln>
								<a:noFill/>
							</a:ln>
							<a:effectLst>
								<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
									<a:srgbClr val="000000">
										<a:alpha val="63000"/>
									</a:srgbClr>
								</a:outerShdw>
							</a:effectLst>
						</c:spPr>
						<c:extLst>
							<c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"
								xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">
								<c16:uniqueId val="{0000000D-83F9-4BC0-8CB9-9896CA7D3BE8}"/>
							</c:ext>
						</c:extLst>
					</c:dPt>
					<c:dPt>
						<c:idx val="10"/>
						<c:invertIfNegative val="0"/>
						<c:bubble3D val="0"/>
						<c:spPr>
							<a:solidFill>
								<a:schemeClr val="accent3"/>
							</a:solidFill>
							<a:ln>
								<a:noFill/>
							</a:ln>
							<a:effectLst>
								<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
									<a:srgbClr val="000000">
										<a:alpha val="63000"/>
									</a:srgbClr>
								</a:outerShdw>
							</a:effectLst>
						</c:spPr>
						<c:extLst>
							<c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"
								xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">
								<c16:uniqueId val="{0000000F-83F9-4BC0-8CB9-9896CA7D3BE8}"/>
							</c:ext>
						</c:extLst>
					</c:dPt>
					<c:dPt>
						<c:idx val="12"/>
						<c:invertIfNegative val="0"/>
						<c:bubble3D val="0"/>
						<c:spPr>
							<a:solidFill>
								<a:schemeClr val="accent2"/>
							</a:solidFill>
							<a:ln>
								<a:noFill/>
							</a:ln>
							<a:effectLst>
								<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
									<a:srgbClr val="000000">
										<a:alpha val="63000"/>
									</a:srgbClr>
								</a:outerShdw>
							</a:effectLst>
						</c:spPr>
						<c:extLst>
							<c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"
								xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">
								<c16:uniqueId val="{00000011-83F9-4BC0-8CB9-9896CA7D3BE8}"/>
							</c:ext>
						</c:extLst>
					</c:dPt>
					<c:dPt>
						<c:idx val="13"/>
						<c:invertIfNegative val="0"/>
						<c:bubble3D val="0"/>
						<c:spPr>
							<a:solidFill>
								<a:schemeClr val="accent3"/>
							</a:solidFill>
							<a:ln>
								<a:noFill/>
							</a:ln>
							<a:effectLst>
								<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
									<a:srgbClr val="000000">
										<a:alpha val="63000"/>
									</a:srgbClr>
								</a:outerShdw>
							</a:effectLst>
						</c:spPr>
						<c:extLst>
							<c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"
								xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">
								<c16:uniqueId val="{00000013-83F9-4BC0-8CB9-9896CA7D3BE8}"/>
							</c:ext>
						</c:extLst>
					</c:dPt>
					<c:cat>
						<c:multiLvlStrRef>
							<c:f>'Value Proposition Analysis'!$B$25:$F$38</c:f>
							<c:multiLvlStrCache>
								<c:ptCount val="14"/>
								<c:lvl>
									<c:pt idx="0">
										<c:v>£</c:v>
									</c:pt>
									<c:pt idx="1">
										<c:v>£</c:v>
									</c:pt>
									<c:pt idx="3">
										<c:v>£</c:v>
									</c:pt>
									<c:pt idx="4">
										<c:v>£</c:v>
									</c:pt>
									<c:pt idx="6">
										<c:v>£</c:v>
									</c:pt>
									<c:pt idx="7">
										<c:v>£</c:v>
									</c:pt>
									<c:pt idx="9">
										<c:v>£</c:v>
									</c:pt>
									<c:pt idx="10">
										<c:v>£</c:v>
									</c:pt>
									<c:pt idx="12">
										<c:v>£</c:v>
									</c:pt>
									<c:pt idx="13">
										<c:v>£</c:v>
									</c:pt>
								</c:lvl>
								<c:lvl>
									<c:pt idx="0">
										<c:v> 1 Year Investment </c:v>
									</c:pt>
									<c:pt idx="1">
										<c:v> 1 Year Return </c:v>
									</c:pt>
									<c:pt idx="3">
										<c:v> 2 Year Investment </c:v>
									</c:pt>
									<c:pt idx="4">
										<c:v> 2 Year Return </c:v>
									</c:pt>
									<c:pt idx="6">
										<c:v> 3 Year Investment </c:v>
									</c:pt>
									<c:pt idx="7">
										<c:v> 3 Year Return </c:v>
									</c:pt>
									<c:pt idx="9">
										<c:v> 4 Year Investment </c:v>
									</c:pt>
									<c:pt idx="10">
										<c:v> 4 Year Return </c:v>
									</c:pt>
									<c:pt idx="12">
										<c:v> 5 Year Investment </c:v>
									</c:pt>
									<c:pt idx="13">
										<c:v> 5 Year Return </c:v>
									</c:pt>
								</c:lvl>
							</c:multiLvlStrCache>
						</c:multiLvlStrRef>
					</c:cat>
					<c:val>
						<c:numRef>
							<c:f>'Value Proposition Analysis'!$G$25:$G$38</c:f>
							<c:numCache>
								<c:formatCode>_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)</c:formatCode>
    
							</c:numCache>
						</c:numRef>
					</c:val>
					<c:extLst>
						<c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"
							xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">
							<c16:uniqueId val="{00000014-83F9-4BC0-8CB9-9896CA7D3BE8}"/>
						</c:ext>
					</c:extLst>
				</c:ser>
				<c:dLbls>
					<c:showLegendKey val="0"/>
					<c:showVal val="0"/>
					<c:showCatName val="0"/>
					<c:showSerName val="0"/>
					<c:showPercent val="0"/>
					<c:showBubbleSize val="0"/>
				</c:dLbls>
				<c:gapWidth val="100"/>
				<c:overlap val="-24"/>
				<c:axId val="945546352"/>
				<c:axId val="945550928"/>
			</c:barChart>
			<c:catAx>
				<c:axId val="945546352"/>
				<c:scaling>
					<c:orientation val="minMax"/>
				</c:scaling>
				<c:delete val="1"/>
				<c:axPos val="b"/>
				<c:numFmt formatCode="General" sourceLinked="1"/>
				<c:majorTickMark val="none"/>
				<c:minorTickMark val="none"/>
				<c:tickLblPos val="nextTo"/>
				<c:crossAx val="945550928"/>
				<c:crosses val="autoZero"/>
				<c:auto val="1"/>
				<c:lblAlgn val="ctr"/>
				<c:lblOffset val="100"/>
				<c:noMultiLvlLbl val="0"/>
			</c:catAx>
			<c:valAx>
				<c:axId val="945550928"/>
				<c:scaling>
					<c:orientation val="minMax"/>
				</c:scaling>
				<c:delete val="0"/>
				<c:axPos val="l"/>
				<c:majorGridlines>
					<c:spPr>
						<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
							<a:solidFill>
								<a:schemeClr val="tx1">
									<a:lumMod val="15000"/>
									<a:lumOff val="85000"/>
								</a:schemeClr>
							</a:solidFill>
							<a:round/>
						</a:ln>
						<a:effectLst/>
					</c:spPr>
				</c:majorGridlines>
				<c:numFmt formatCode="_(* #,##0_);_(* \(#,##0\);_(* &quot;-&quot;??_);_(@_)" sourceLinked="1"/>
				<c:majorTickMark val="none"/>
				<c:minorTickMark val="none"/>
				<c:tickLblPos val="nextTo"/>
				<c:spPr>
					<a:noFill/>
					<a:ln>
						<a:noFill/>
					</a:ln>
					<a:effectLst/>
				</c:spPr>
				<c:txPr>
					<a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
					<a:lstStyle/>
					<a:p>
						<a:pPr>
							<a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">
								<a:solidFill>
									<a:schemeClr val="tx1">
										<a:lumMod val="65000"/>
										<a:lumOff val="35000"/>
									</a:schemeClr>
								</a:solidFill>
								<a:latin typeface="+mn-lt"/>
								<a:ea typeface="+mn-ea"/>
								<a:cs typeface="+mn-cs"/>
							</a:defRPr>
						</a:pPr>
						<a:endParaRPr lang="tr-TR"/>
					</a:p>
				</c:txPr>
				<c:crossAx val="945546352"/>
				<c:crosses val="autoZero"/>
				<c:crossBetween val="between"/>
			</c:valAx>
			<c:spPr>
				<a:noFill/>
				<a:ln>
					<a:noFill/>
				</a:ln>
				<a:effectLst/>
			</c:spPr>
		</c:plotArea>
		<c:plotVisOnly val="1"/>
		<c:dispBlanksAs val="gap"/>
		<c:extLst>
			<c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}"
				xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart">
				<c16r3:dataDisplayOptions16>
					<c16r3:dispNaAsBlank val="1"/>
				</c16r3:dataDisplayOptions16>
			</c:ext>
		</c:extLst>
		<c:showDLblsOverMax val="0"/>
	</c:chart>
	<c:spPr>
		<a:noFill/>
		<a:ln>
			<a:noFill/>
		</a:ln>
		<a:effectLst/>
	</c:spPr>
	<c:txPr>
		<a:bodyPr/>
		<a:lstStyle/>
		<a:p>
			<a:pPr>
				<a:defRPr/>
			</a:pPr>
			<a:endParaRPr lang="tr-TR"/>
		</a:p>
	</c:txPr>
	<c:externalData r:id="rId3">
		<c:autoUpdate val="0"/>
	</c:externalData>
</c:chartSpace>
'''
    return xml_content

def update_zip_with_new_xml(zip_path, output_zip_path):
    temp_dir = 'temp_zip'
    os.makedirs(temp_dir, exist_ok=True)

    # ZIP dosyasını çıkar
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    # Yeni XML içeriğini oluştur
    chart_xml_content = create_chart_xml()

    # Yeni içeriği chart1.xml olarak kaydet
    new_xml_path = os.path.join(temp_dir, 'ppt', 'charts', 'chart1.xml')
    with open(new_xml_path, 'w', encoding='utf-8') as xml_file:
        xml_file.write(chart_xml_content)

    # <c:numCache> içerisine yeni verileri ekle
    with open(new_xml_path, 'r+', encoding='utf-8') as xml_file:
        lines = xml_file.readlines()
        for index, line in enumerate(lines):
            if '<c:numCache>' in line:
                # Yeni verileri ekle
                lines.insert(index + 1, '    <c:ptCount val="14"/>\n')
                lines.insert(index + 2, '    <c:pt idx="0"><c:v>13000</c:v></c:pt>\n')
                lines.insert(index + 3, '    <c:pt idx="1"><c:v>16430.547826086971</c:v></c:pt>\n')

                lines.insert(index + 2, '    <c:pt idx="3"><c:v>13000</c:v></c:pt>\n')
                lines.insert(index + 3, '    <c:pt idx="4"><c:v>16430.547826086971</c:v></c:pt>\n')

                lines.insert(index + 2, '    <c:pt idx="6"><c:v>13000</c:v></c:pt>\n')
                lines.insert(index + 3, '    <c:pt idx="7"><c:v>16430.547826086971</c:v></c:pt>\n')

                lines.insert(index + 2, '    <c:pt idx="9"><c:v>13000</c:v></c:pt>\n')
                lines.insert(index + 3, '    <c:pt idx="10"><c:v>16430.547826086971</c:v></c:pt>\n')

                lines.insert(index + 2, '    <c:pt idx="12"><c:v>13000</c:v></c:pt>\n')
                lines.insert(index + 3, '    <c:pt idx="13"><c:v>16430.547826086971</c:v></c:pt>\n')
                break
        xml_file.seek(0)
        xml_file.writelines(lines)

    # Güncellenmiş dosyaları yeni ZIP dosyası olarak kaydet
    with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
        for foldername, subfolders, filenames in os.walk(temp_dir):
            for filename in filenames:
                filepath = os.path.join(foldername, filename)
                arcname = os.path.relpath(filepath, temp_dir)
                zip_ref.write(filepath, arcname)

    # Geçici dizini temizle
    shutil.rmtree(temp_dir)

# Kullanım
zip_path = 'template.zip'  # Güncellemek istediğin template.zip
output_zip_path = 'template.zip'  # Çıkış dosyasının adı

# Yeni XML dosyasını oluştur ve ZIP dosyasını güncelle
update_zip_with_new_xml(zip_path, output_zip_path)
