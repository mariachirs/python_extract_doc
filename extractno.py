REQUIRED_TYPES = [
    "Micro-ordinateurs",
    "Bases de données",
    "Logiciels d’exploitation",
    "Outils centraux",
    "Langages de programmation",
    "Logiciels de navigation Internet"
]

def extract_technologies_by_required_types(doc_path, required_types):
    doc = Document(doc_path)
    tech_entries = []
    current_type = None
    found_types = set()

    for table in doc.tables:
        first_row = table.rows[0].cells
        headers = [cell.text.strip() for cell in first_row]

        # Find the right table with the expected headers
        if "Technologies" in headers and "Mois" in headers:
            # Process every row in this table
            for row in table.rows[1:]:
                cells = row.cells
                texts = [cell.text.strip() for cell in cells]
                if all(not text for text in texts):
                    continue

                for i in range(0, len(cells), 2):
                    tech = cells[i].text.strip() if i < len(cells) else ''
                    mois = cells[i + 1].text.strip() if i + 1 < len(cells) else ''

                    # Detect section titles like "Outils centraux:"
                    if tech.endswith(":") and not mois:
                        section = tech.replace(":", "").strip()
                        if section in required_types:
                            current_type = section
                            found_types.add(section)
                        else:
                            current_type = None
                        continue

                    elif mois.endswith(":") and not tech:
                        section = mois.replace(":", "").strip()
                        if section in required_types:
                            current_type = section
                            found_types.add(section)
                        else:
                            current_type = None
                        continue

                    if tech and mois and current_type in required_types:
                        try:
                            mois_int = int(mois)
                            tech_entries.append({
                                "technology": tech,
                                "mois": mois_int,
                                "type": current_type
                            })
                        except ValueError:
                            continue

    missing_types = set(required_types) - found_types
    return tech_entries, missing_types

# Example usage:
doc_path = "MARTINS_Roni_CV_24-02-13.docx"  # Change this path to your DOCX file
entries, missing = extract_technologies_by_required_types(doc_path, REQUIRED_TYPES)



///////////////////////////////////////


def extract_all_technology_rows(doc_path):
    doc = Document(doc_path)
    all_rows = []

    for table in doc.tables:
        # Detect valid table by headers
        first_row = table.rows[0].cells
        headers = [cell.text.strip().replace('\xa0', ' ') for cell in first_row]
        if "Technologies" in headers and "Mois" in headers:
            for row in table.rows[1:]:  # skip header
                cells = [cell.text.strip().replace('\xa0', ' ') for cell in row.cells]
                # Ensure each row has 4 items
                while len(cells) < 4:
                    cells.append('')
                all_rows.append(cells[:4])  # Only take the first 4 columns

    return all_rows
