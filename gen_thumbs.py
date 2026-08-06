"""
Generate slide thumbnails: PPTX -> PDF (LibreOffice) -> PNG (PyMuPDF)
Usage: python gen_thumbs.py <input.pptx> <output_dir>
"""
import subprocess, os, sys, tempfile
import fitz  # PyMuPDF

SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"

def gen(pptx_path, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    
    with tempfile.TemporaryDirectory() as tmp:
        # Step 1: PPTX → PDF via LibreOffice
        print(f"[1/2] Converting {os.path.basename(pptx_path)} to PDF...")
        subprocess.run([
            SOFFICE, "--headless", "--norestore",
            "--convert-to", "pdf", "--outdir", tmp, pptx_path
        ], capture_output=True, text=True, timeout=180)
        
        pdfs = [f for f in os.listdir(tmp) if f.endswith('.pdf')]
        if not pdfs:
            print("ERROR: No PDF generated!")
            return
        
        pdf_path = os.path.join(tmp, pdfs[0])
        
        # Step 2: PDF → PNGs via PyMuPDF
        print(f"[2/2] Rendering slides to PNG...")
        doc = fitz.open(pdf_path)
        for i in range(len(doc)):
            page = doc[i]
            # Render at 2x for crisp quality
            mat = fitz.Matrix(2.0, 2.0)
            pix = page.get_pixmap(matrix=mat)
            out = os.path.join(out_dir, f"slide_{i+1}.png")
            pix.save(out)
            print(f"  slide_{i+1}.png ({pix.width}x{pix.height})")
        count = len(doc)
        doc.close()
        
        print(f"\nDone! {count} slides → {out_dir}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python gen_thumbs.py <input.pptx> <output_dir>")
        sys.exit(1)
    gen(sys.argv[1], sys.argv[2])
