from flask import Flask, render_template, request, send_file
from docx import Document
from io import BytesIO
import base64
from PIL import Image
from docx2pdf import convert
import shutil
import subprocess
import io, datetime, os, re, copy, traceback, logging

logging.basicConfig(level=logging.INFO)

app = Flask(__name__)


@app.route('/download/<child_folder>/<filename>')
def download_file(child_folder, filename):
    # basic validation to prevent directory traversal
    if '/' in child_folder or '..' in child_folder or '/' in filename or '..' in filename:
        return "Invalid path", 400
    safe_base = os.path.join(os.getcwd(), 'temp')
    file_path = os.path.join(safe_base, child_folder, filename)
    if not os.path.exists(file_path):
        return "File not found", 404
    return send_file(file_path, as_attachment=True)

# Note: HTML moved to templates/form.html and static/styles.css

@app.route('/')
def form():
    return render_template('form.html')

@app.route('/generate', methods=['POST'])
def generate_docx():
    data = request.form.to_dict()

    # Generare automată număr și dată contract
    today = datetime.datetime.now()
    data['data_contract'] = today.strftime('%d.%m.%Y')
    data['numar_contract'] = f"ARIEL-{today.strftime('%Y%m%d-%H%M%S')}"

    # Creăm un folder temporar sigur în proiect per copil
    base_temp = os.path.join(os.getcwd(), 'temp')
    os.makedirs(base_temp, exist_ok=True)

    # Construim numele folderului din nume + prenume copil și îl sanitizăm
    child_raw = f"{data.get('nume_copil','').strip()}_{data.get('prenume_copil','').strip()}".strip('_ ')
    def _sanitize(s):
        s = s.strip()
        s = re.sub(r'[^A-Za-z0-9_-]', '_', s)
        s = re.sub(r'_+', '_', s)
        return s or 'copil'

    child_name = _sanitize(child_raw)

    # Dacă folderul există deja, adăugăm sufix _1, _2, ...
    candidate = os.path.join(base_temp, child_name)
    count = 1
    while os.path.exists(candidate):
        candidate = os.path.join(base_temp, f"{child_name}_{count}")
        count += 1

    temp_dir = candidate
    os.makedirs(temp_dir, exist_ok=True)

    # Move any stray files from the base temp folder into this child's folder
    try:
        for entry in os.listdir(base_temp):
            entry_path = os.path.join(base_temp, entry)
            # skip directories (including other child folders)
            if os.path.isfile(entry_path):
                # move files (e.g., leftover .docx/.pdf/.png) into the child folder
                shutil.move(entry_path, os.path.join(temp_dir, entry))
    except Exception:
        # don't fail the whole request if cleanup/move can't run
        pass
        
    signature_path = os.path.join(temp_dir, 'signature.png')

    # Salvăm semnătura ca imagine
    signature_data = data['signature_data'].split(',')[1]
    signature_bytes = base64.b64decode(signature_data)
    signature_img = Image.open(io.BytesIO(signature_bytes))
    signature_img.save(signature_path)

    # Template-uri DOCX (use project-relative paths)
    base_dir = os.path.dirname(os.path.abspath(__file__))
    templates_dir = os.path.join(base_dir, 'templates')
    contracts = [
        (os.path.join(templates_dir, 'educational.docx'), "contract_educational_completat.docx"),
        (os.path.join(templates_dir, 'catering.docx'), "contract_catering_completat.docx")
    ]

    generated_pdfs = []
    conversion_errors = {}
    # compile a regex to catch {{semnatura}} with optional spaces and case-insensitive
    sig_pattern = re.compile(r"\{\{\s*semnatura\s*\}\}", flags=re.IGNORECASE)
    for template_path, output_name in contracts:
        doc = Document(template_path)
        signature_inserted = False

        def replace_placeholders_in_paragraph(p):
            nonlocal signature_inserted
            # replace normal placeholders from form data
            for key, val in data.items():
                if key != 'signature_data' and f'{{{{{key}}}}}' in p.text:
                    p.text = p.text.replace(f'{{{{{key}}}}}', val)

            # check for signature placeholder variants
            if sig_pattern.search(p.text):
                # remove placeholder
                p.text = sig_pattern.sub('', p.text)
                try:
                    # add picture which creates a new paragraph with the image
                    doc.add_picture(signature_path)
                    pic_para = doc.paragraphs[-1]
                    # get the drawing element from the picture run
                    pic_run = pic_para.runs[0] if pic_para.runs else None
                    drawing_elem = None
                    if pic_run is not None:
                        for child in pic_run._r:
                            # look for drawing element (namespace-aware)
                            if 'drawing' in child.tag:
                                drawing_elem = child
                                break
                    if drawing_elem is not None:
                        # create a new run in the target paragraph and insert the drawing
                        new_run = p.add_run()
                        new_run._r.append(copy.deepcopy(drawing_elem))
                        # move the new run to the start of the paragraph so image appears before text
                        p._p.insert(0, new_run._r)
                        # remove the temporary picture paragraph
                        try:
                            pic_para._p.getparent().remove(pic_para._p)
                        except Exception:
                            pass
                        signature_inserted = True
                except Exception:
                    pass

        # process top-level paragraphs
        for p in doc.paragraphs:
            replace_placeholders_in_paragraph(p)

        # process paragraphs inside table cells as well
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        replace_placeholders_in_paragraph(p)

        # If no {{semnatura}} was present/inserted, add signature block at end
        if not signature_inserted:
            doc.add_paragraph(f"\nSemnătură părinte:\n{data.get('nume_mama', '')} / {data.get('nume_tata', '')}")
            try:
                doc.add_picture(signature_path)
            except Exception:
                pass
        output_path = os.path.join(temp_dir, output_name)
        doc.save(output_path)

        # Convert DOCX to PDF (try docx2pdf, then LibreOffice/soffice fallback)
        pdf_name = os.path.splitext(output_name)[0] + '.pdf'
        pdf_path = os.path.join(temp_dir, pdf_name)

        def try_convert_to_pdf(input_docx, output_pdf):
            # Try docx2pdf (MS Word COM on Windows)
            try:
                convert(input_docx, output_pdf)
                return True, None
            except Exception as e:
                err = traceback.format_exc()
                logging.info(f"docx2pdf failed for {input_docx}: {e}")
                # continue to fallback

            # Try LibreOffice/soffice headless conversion
            soffice = shutil.which('soffice') or shutil.which('libreoffice')
            if soffice:
                try:
                    outdir = os.path.dirname(output_pdf)
                    subprocess.run([soffice, '--headless', '--convert-to', 'pdf', '--outdir', outdir, input_docx], check=True)
                    if os.path.exists(output_pdf):
                        return True, None
                    else:
                        return False, 'LibreOffice conversion did not produce output file.'
                except Exception as e:
                    err = traceback.format_exc()
                    logging.info(f"LibreOffice conversion failed for {input_docx}: {e}")
                    return False, err

            # no converter available
            return False, 'No available converter (docx2pdf/soffice)'

        converted, err = try_convert_to_pdf(output_path, pdf_path)
        if converted:
            generated_pdfs.append(pdf_name)
        else:
            # conversion failed; record error but DO NOT append DOCX fallback
            conversion_errors[output_name] = err
            logging.error(f"Conversion failed for {output_name}: {err}")

    # Prepare download items (only PDFs)
    final_child_folder = os.path.basename(temp_dir)
    pdf_files = generated_pdfs
    download_items = [(name, f"/download/{final_child_folder}/{name}") for name in pdf_files]

    if conversion_errors:
        return render_template('download.html', files=download_items, failed=list(conversion_errors.keys()), errors=conversion_errors)

    return render_template('download.html', files=download_items, failed=None, errors=None)

if __name__ == '__main__':
    app.run(debug=True)