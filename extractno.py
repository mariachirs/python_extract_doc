from docx import Document

def extract_technologies_with_type(doc_path):
    doc = Document(doc_path)
    tech_entries = []
    current_type = None

    for table in doc.tables:
        # Check for header in the first row
        first_row_texts = [cell.text.strip() for cell in table.rows[0].cells]
        if "Technologies" in first_row_texts and "Mois" in first_row_texts:
            for row in table.rows[1:]:
                cells = row.cells
                for i in range(0, len(cells), 2):
                    tech_raw = cells[i].text.strip().replace('\xa0', ' ')
                    mois_raw = cells[i+1].text.strip() if i+1 < len(cells) else ''

                    # Detect and update current type
                    if tech_raw.endswith(":") and not mois_raw:
                        current_type = tech_raw.replace(":", "").strip()
                        continue

                    # Skip empty entries
                    if not tech_raw or not mois_raw:
                        continue

                    try:
                        mois_int = int(mois_raw)
                        tech_entries.append({
                            "technology": tech_raw,
                            "mois": mois_int,
                            "type": current_type
                        })
                    except ValueError:
                        continue  # skip rows where "mois" is not a number

    return {
        entry["technology"]: {
            "mois": entry["mois"],
            "type": entry["type"]
        }
        for entry in tech_entries
    }

# Usage
tech_dict = extract_technologies_with_type("MARTINS_Roni_CV_24-02-13.docx")

from pprint import pprint
pprint(tech_dict)
