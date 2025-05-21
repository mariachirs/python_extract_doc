from docx.document import Document

def extract_certifications_from_doc(doc: Document, target_header="Certifications"):
    """
    Extract certification entries from a docx.Document.
    Each certification is assumed to span 4 lines:
    Title, Year, Organization, Location.
    Stops when a known header is found.
    """
    # Hardcoded headers where we stop capturing
    stop_headers = [
        "Résumé des interventions",
        "Perfectionnement",
        "Langues parlées, écrites",
        "Formation académique",
        "Principaux domaines"
    ]
    stop_headers_lower = [h.lower() for h in stop_headers]
    target_header_lower = target_header.strip().lower()

    capturing = False
    collected = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        lower_text = text.lower()

        if not capturing and lower_text == target_header_lower:
            capturing = True
            continue

        if capturing:
            if lower_text in stop_headers_lower:
                break
            collected.append(text)

    # Group every 4 lines into one certification
    certifications = []
    for i in range(0, len(collected), 4):
        group = collected[i:i+4]
        if len(group) == 4:
            certifications.append({
                "Title": group[0],
                "Year": group[1],
                "Organization": group[2],
                "Location": group[3]
            })

    return certifications

from docx import Document

doc = Document("CV-Gabarit-LGS-2023.docx")
certs = extract_certifications_from_doc(doc)

for cert in certs:
    print(cert)
