from flask import Flask, request, jsonify, send_file
import zipfile, shutil, os, re, io, base64
from datetime import datetime, timezone, timedelta

app = Flask(__name__)

# Mapeo WMO → imagen del PPTX
def get_imagen_clima(wmo):
    if   wmo == 0:   return "image14.png"  # Sol
    elif wmo <= 2:   return "image4.png"   # Parcialmente nublado
    elif wmo <= 55:  return "image18.png"  # Nublado
    elif wmo <= 67:  return "image16.png"  # Lluvia
    elif wmo <= 77:  return "image18.png"  # Nevada
    elif wmo <= 82:  return "image12.png"  # Sol con lluvia
    else:            return "image2.png"   # Tormenta

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok"})

@app.route('/generar_clima', methods=['POST'])
def generar_clima():
    try:
        data = request.json
        
        # Datos del clima
        temp       = data.get('temp', 15)
        sensacion  = data.get('sensacion', 13)
        humedad    = data.get('humedad', 60)
        viento     = data.get('viento', 10)
        wmo        = data.get('wmo', 3)
        wmo_pron   = data.get('wmo_pron', [3, 3, 3])
        slide_num  = data.get('slide_num', 3)
        dia_nombre = data.get('dia_nombre', 'Lunes')
        fecha_str  = data.get('fecha_str', '13 de julio')
        prox_dias  = data.get('prox_dias', ['Martes', 'Miércoles', 'Jueves'])
        maximas_pron = data.get('maximas_pron', [15, 16, 17])
        minimas_pron = data.get('minimas_pron', [5, 6, 7])
        probs_pron   = data.get('probs_pron', [10, 20, 10])
        
        # PPTX base64
        pptx_b64 = data.get('pptx_b64', '')
        if not pptx_b64:
            return jsonify({"error": "pptx_b64 requerido"}), 400
        
        pptx_bytes = base64.b64decode(pptx_b64)
        
        # Desempaquetar PPTX
        dst = '/tmp/clima_build'
        if os.path.exists(dst):
            shutil.rmtree(dst)
        with zipfile.ZipFile(io.BytesIO(pptx_bytes), 'r') as z:
            z.extractall(dst)
        
        # Editar slide correcta
        slide_path = f"{dst}/ppt/slides/slide{slide_num}.xml"
        rels_path  = f"{dst}/ppt/slides/_rels/slide{slide_num}.xml.rels"
        
        with open(slide_path, 'r', encoding='utf-8') as f:
            xml = f.read()
        
        def rep(xml, old, new, n=1):
            return xml.replace(f'>{old}<', f'>{new}<', n)
        
        # Reemplazar textos
        xml = rep(xml, "Miércoles",       dia_nombre)
        xml = rep(xml, "15 de noviembre", fecha_str)
        xml = rep(xml, "18º",             f"{temp}º")
        xml = rep(xml, "Jueves",          prox_dias[0])
        xml = rep(xml, "Viernes",         prox_dias[1])
        xml = rep(xml, "Sábado",          prox_dias[2])
        xml = rep(xml, "##SENS##",        f"{sensacion}° C")
        xml = rep(xml, "##MAX1##",        f"{maximas_pron[0]}° C")
        xml = rep(xml, "##MAX2##",        f"{maximas_pron[1]}° C")
        xml = rep(xml, "##MAX3##",        f"{maximas_pron[2]}° C")
        xml = rep(xml, "##VTO0##",        f"{viento} Km")
        xml = rep(xml, "##VTO1##",        f"{viento} Km")
        xml = rep(xml, "##VTO2##",        f"{viento} Km")
        xml = rep(xml, "##VTO3##",        f"{viento} Km")
        xml = rep(xml, "##HUM0##",        f"{humedad}%")
        xml = rep(xml, "##PROB1##",       f"{probs_pron[0]}%")
        xml = rep(xml, "##PROB2##",       f"{probs_pron[1]}%")
        xml = rep(xml, "##PROB3##",       f"{probs_pron[2]}%")
        
        with open(slide_path, 'w', encoding='utf-8') as f:
            f.write(xml)
        
        # Reemplazar íconos de los cards en el XML
        # Leer rels para mapear rIds
        with open(rels_path, 'r') as f:
            rels = f.read()
        
        rid_map = {}
        for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels):
            rid_map[m.group(1)] = os.path.basename(m.group(2))
        
        from xml.etree import ElementTree as ET
        PNS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
        ANS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        
        tree = ET.fromstring(xml)
        
        # Encontrar Freeform 18, Freeform 32/51, Freeform 19 (cards)
        card_freeforms = ['Freeform 18', 'Freeform 32', 'Freeform 51', 'Freeform 19']
        card_shapes = []
        
        for sp in tree.iter(f'{{{PNS}}}sp'):
            nvpr = sp.find(f'.//{{{PNS}}}cNvPr')
            if nvpr is None: continue
            name = nvpr.get('name', '')
            if name in card_freeforms:
                blip = sp.find(f'.//{{{ANS}}}blip')
                if blip is not None:
                    rid = blip.get(f'{{{RNS}}}embed', '')
                    card_shapes.append({'name': name, 'rid': rid, 'blip': blip})
        
        # Ordenar: Freeform 18 (card1), Freeform 32 o 51 (card2), Freeform 19 (card3)
        card_order = {}
        for cs in card_shapes:
            if cs['name'] == 'Freeform 18':  card_order[0] = cs
            elif cs['name'] in ['Freeform 32', 'Freeform 51']: card_order[1] = cs
            elif cs['name'] == 'Freeform 19': card_order[2] = cs
        
        # Para cada card, actualizar el rId para que apunte a la imagen correcta
        nueva_rels = rels
        max_rid = max([int(r) for r in re.findall(r'rId(\d+)', rels)]) if re.findall(r'rId(\d+)', rels) else 0
        
        nuevo_xml = xml
        for idx, wmo_dia in enumerate(wmo_pron[:3]):
            nueva_img = get_imagen_clima(wmo_dia)
            if idx not in card_order: continue
            cs = card_order[idx]
            img_actual = rid_map.get(cs['rid'], '')
            
            if img_actual == nueva_img:
                continue
            
            # Agregar nuevo rId para la imagen
            max_rid += 1
            new_rid = f"rId{max_rid}"
            nueva_rels = nueva_rels.replace(
                '</Relationships>',
                f'<Relationship Id="{new_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{nueva_img}"/>\n</Relationships>'
            )
            
            # Reemplazar el embed en el XML dentro del shape específico
            idx_start = nuevo_xml.find(f'name="{cs["name"]}"')
            if idx_start == -1: continue
            idx_end = nuevo_xml.find('</p:sp>', idx_start)
            sp_xml = nuevo_xml[idx_start:idx_end]
            new_sp = sp_xml.replace(f'r:embed="{cs["rid"]}"', f'r:embed="{new_rid}"', 1)
            
            # Hacer el ícono cuadrado — reemplazar <a:ext cx="X" cy="Y"/> con cuadrado
            import re as re2
            def hacer_cuadrado(m):
                cx_v = int(m.group(1))
                cy_v = int(m.group(2))
                min_v = min(cx_v, cy_v)
                return f'<a:ext cx="{min_v}" cy="{min_v}"/>'
            new_sp = re2.sub(r'<a:ext cx="(\d+)" cy="(\d+)"/>', hacer_cuadrado, new_sp)
            
            nuevo_xml = nuevo_xml[:idx_start] + new_sp + nuevo_xml[idx_end:]
        
        with open(slide_path, 'w', encoding='utf-8') as f:
            f.write(nuevo_xml)
        with open(rels_path, 'w') as f:
            f.write(nueva_rels)
        
        # Dejar solo slide elegida
        pres_path = f"{dst}/ppt/presentation.xml"
        pres_rels  = f"{dst}/ppt/_rels/presentation.xml.rels"
        with open(pres_path, 'r') as f: pres_xml = f.read()
        with open(pres_rels, 'r')  as f: pr_rels  = f.read()
        
        match = re.search(rf'Id="(rId\d+)"[^>]*Target="slides/slide{slide_num}\.xml"', pr_rels)
        if not match:
            match = re.search(rf'Target="slides/slide{slide_num}\.xml"[^>]*Id="(rId\d+)"', pr_rels)
        
        if match:
            our_rid = match.group(1)
            all_sldids = re.findall(r'<p:sldId\b[^/]*/>', pres_xml)
            our_sldid  = next((s for s in all_sldids if our_rid in s), None)
            if our_sldid:
                pres_xml = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>',
                                  f'<p:sldIdLst>{our_sldid}</p:sldIdLst>',
                                  pres_xml, flags=re.DOTALL)
                with open(pres_path, 'w') as f: f.write(pres_xml)
        
        # Reempaquetar
        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED) as z:
            for root2, dirs, files in os.walk(dst):
                for file in files:
                    fp = os.path.join(root2, file)
                    z.write(fp, os.path.relpath(fp, dst))
        
        out_buf.seek(0)
        return send_file(
            out_buf,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            as_attachment=True,
            download_name='clima_editado.pptx'
        )
    
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
