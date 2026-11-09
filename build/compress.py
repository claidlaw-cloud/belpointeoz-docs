# Compresses the deck PDF: downsamples images to max 1400px and re-encodes as JPEG q82.
import sys, io, pikepdf
from PIL import Image
src, dst = sys.argv[1], sys.argv[2]
pdf = pikepdf.open(src)
for page in pdf.pages:
    for name, raw in list(page.images.items()):
        if "/SMask" in raw or "/Mask" in raw:
            continue  # keep transparent images (logos) untouched
        try:
            pim = pikepdf.PdfImage(raw); img = pim.as_pil_image()
        except Exception:
            continue
        if img.mode not in ("RGB", "L"): img = img.convert("RGB")
        w, h = img.size; m = max(w, h)
        if m > 1400: img = img.resize((round(w * 1400 / m), round(h * 1400 / m)), Image.LANCZOS)
        buf = io.BytesIO(); img.save(buf, "JPEG", quality=82, optimize=True)
        if buf.tell() < len(raw.read_raw_bytes()):
            raw.write(buf.getvalue(), filter=pikepdf.Name("/DCTDecode"))
            raw.Width, raw.Height = img.size
            raw.ColorSpace = pikepdf.Name("/DeviceRGB") if img.mode == "RGB" else pikepdf.Name("/DeviceGray")
            raw.BitsPerComponent = 8
            for k in ("/Decode", "/DecodeParms"):
                if k in raw: del raw[k]
pdf.save(dst, compress_streams=True, object_stream_mode=pikepdf.ObjectStreamMode.generate)
