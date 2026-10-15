# AZ PDF LOVER — Render backend edition

This is a deployable full-stack version of AZ PDF LOVER. The existing responsive single-page frontend is served by Flask and sends file-processing jobs to a Python backend. The backend uses PyMuPDF, pikepdf, Pillow, Tesseract OCR, LibreOffice, and Ghostscript.

## What the backend supports

- Merge PDFs; split/extract page ranges; delete pages; reorder pages; rotate pages.
- Compress/re-save PDFs, set metadata, crop page views, repair/re-save parseable PDFs.
- Add text annotations, text signatures, page numbers, text watermarks, and image overlays.
- Password-protect PDFs with AES-256 encryption and remove encryption when the supplied password is valid. It does not bypass unknown passwords.
- PDF to JPG (ZIP); JPG to PDF; PDF text to DOCX/TXT; PDF text to XLSX/CSV; PDF pages to PPTX slides; DOCX/PPTX/XLSX to PDF via LibreOffice.
- OCR selected PDF pages to TXT in English, Urdu, Spanish, French, or German (language packs installed in Docker).
- Compare two PDFs using extracted text and produce a diff report.

## Important fidelity notes

A backend makes many more operations possible, but no tool can guarantee perfect results for every PDF. PDF-to-Word and PDF-to-Excel export selectable text rather than reconstructing every original layout/table. PDF-to-PowerPoint uses one page image per slide, so the slide content is not editable. OCR accuracy depends on scan quality and language. LibreOffice conversion may differ from Microsoft Office on complex documents. Repair only works when the underlying PDF parser can recover the file. PDF password removal requires a valid password and authorization.

## Deploy on Render

1. Download and extract `az_pdf_lover_render_bundle.zip`.
2. Create a new GitHub repository and upload **all extracted files**, keeping `Dockerfile`, `render.yaml`, `app.py`, `requirements.txt`, and the `static` folder at the repository root.
3. Open Render and choose **New + → Blueprint** (or create a Docker web service connected to the repo). If using Blueprint, Render reads `render.yaml`.
4. Select the repository and deploy. The Docker build installs LibreOffice, Ghostscript, Tesseract and the Python dependencies. First deploy can take several minutes.
5. When the service is live, open `https://YOUR-SERVICE.onrender.com/health`. It should return `{"status":"ok","service":"AZ PDF LOVER API"}`. Then open the root URL and test with non-sensitive sample files.
6. For a manual Docker Web Service, set runtime to Docker, health check path `/health`, and leave the start command blank (the Dockerfile provides it). The service listens on Render's `$PORT`.

## Limits and privacy

- Default limits: 100 MB per file and 220 MB per request. Change `MAX_FILE_MB` / `MAX_REQUEST_MB` environment variables if your plan can handle larger files.
- Render's filesystem is ephemeral; this app also deletes each request's temporary directory in a `finally` block. Do not add persistent storage unless you deliberately need it.
- Files are uploaded to your server for processing. Publish a privacy policy that reflects your actual deployment, do not claim browser-only/local processing, and avoid uploading sensitive documents to a service you do not control.
- A free or low-memory instance may time out or run out of memory on large scans/Office files. For production use, add authentication, rate limiting, request quotas, antivirus scanning, and a queue for long-running jobs. Never expose this endpoint as an unrestricted public conversion service without abuse controls.
- Do not commit secrets to GitHub. No API keys are required for this build.

## Local run (optional)

With Docker installed:

```bash
docker build -t az-pdf-lover .
docker run --rm -p 10000:10000 az-pdf-lover
```

Then open `http://localhost:10000`.
