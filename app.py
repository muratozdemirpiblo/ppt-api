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

def update_zip_with_new_xml(zip_path, output_zip_path, year1invest, year1return, year2invest, year2return,
                             year3invest, year3return, year4invest, year4return, year5invest, year5return):
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

    with open(new_xml_path, 'r+', encoding='utf-8') as xml_file:
        lines = xml_file.readlines()
        for index, line in enumerate(lines):
            if '<c:numCache>' in line:
                # Yeni verileri ekle
                
                break
        xml_file.seek(0)
        xml_file.writelines(lines)


    # <c:numCache> içerisine yeni verileri ekle
    with open(new_xml_path, 'r+', encoding='utf-8') as xml_file:
        lines = xml_file.readlines()
        for index, line in enumerate(lines):
            if '<c:numCache>' in line:
                # Yeni verileri ekle
                lines.insert(index + 1, '    <c:ptCount val="10"/>\n')
                lines.insert(index + 2, '    <c:pt idx="0"><c:v>{}</c:v></c:pt>\n'.format(year1invest))
                lines.insert(index + 3, '    <c:pt idx="1"><c:v>{}</c:v></c:pt>\n'.format(year1return))
                lines.insert(index + 4, '    <c:pt idx="2"><c:v>{}</c:v></c:pt>\n'.format(year2invest))
                lines.insert(index + 5, '    <c:pt idx="3"><c:v>{}</c:v></c:pt>\n'.format(year2return))
                lines.insert(index + 6, '    <c:pt idx="4"><c:v>{}</c:v></c:pt>\n'.format(year3invest))
                lines.insert(index + 7, '    <c:pt idx="5"><c:v>{}</c:v></c:pt>\n'.format(year3return))
                lines.insert(index + 8, '    <c:pt idx="6"><c:v>{}</c:v></c:pt>\n'.format(year4invest))
                lines.insert(index + 9, '    <c:pt idx="7"><c:v>{}</c:v></c:pt>\n'.format(year4return))
                lines.insert(index + 10, '    <c:pt idx="8"><c:v>{}</c:v></c:pt>\n'.format(year5invest))
                lines.insert(index + 11, '    <c:pt idx="9"><c:v>{}</c:v></c:pt>\n'.format(year5return))
                break
        xml_file.seek(0)
        xml_file.writelines(lines)

    # Güncellenmiş dosyaları yeni ZIP dosyası olarak kaydet
    with zipfile.ZipFile('template.zip', 'w', zipfile.ZIP_DEFLATED) as zip_ref:
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
                                year5return='0'):
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
            if 'valhours' in elem.text:
                elem.text = elem.text.replace('valhours', valhours)
        zip_path = 'template.zip'  # Güncellemek istediğin template.zip
        output_zip_path = 'template.zip'  # Çıkış dosyasının adı

        # Yeni XML dosyasını oluştur ve ZIP dosyasını güncelle
        update_zip_with_new_xml(zip_path, output_zip_path,year1invest=year1invest,
                                year1return=year1return,year2invest=year2invest,year2return=year2return,year3invest=year3invest,
                                year3return=year3return,year4invest=year4invest,year4return=year4return,year5invest=year5invest,
                                year5return=year5return)

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
    valhours = data.get('valhours')
    year1total = data.get('year1total')
    year1invest = data.get('year1invest')

    year2otal = data.get('year2total')
    year2invest = data.get('year2invest')

    year3total = data.get('year3total')
    year3invest = data.get('year3invest')

    year4total = data.get('year4total')
    year4invest = data.get('year4invest')

    year5total = data.get('year5total')
    year5invest = data.get('year5invest')
    
    
    if not client_name:
        return "Error: 'client_name' parameter is required", 400





    zip_path = r"template.zip"  # Tam dosya yolunu girin
    output_pptx_path = r"output.pptx"  # Çıkış dosyasının yolunu belirtin

    modify_slide_xml_and_image(zip_path, output_pptx_path,client_name,itfinance,rpo,poa,cip,mspi,valmsl,valfqmr,valdcap,
                               valcifw,valoem,valbnft,valnpvv,valacd,valroi,valinvestment,valmonths,valhours,
                               year1invest=year1invest,
                                year1return=year1total,year2invest=year2invest,year2return=year2otal,year3invest=year3invest,
                                year3return=year3total,year4invest=year4invest,year4return=year4total,year5invest=year5invest,
                                year5return=year5total)
    
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




if __name__ == "__main__":
    app.run(host='0.0.0.0', port=8080)
