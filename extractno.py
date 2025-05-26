def extract_technologies_with_type(doc_path):
    doc = Document(doc_path)
    tech_entries = []
    current_type = None

    for table in doc.tables:
        # Check if it's a relevant table (has Mois headers)
        first_row = table.rows[0].cells
        headers = [cell.text.strip() for cell in first_row]
        if "Technologies" in headers and "Mois" in headers:
            for row in table.rows[1:]:
                cells = row.cells
                for i in range(0, len(cells), 2):
                    if i + 1 < len(cells):
                        tech = cells[i].text.strip()
                        mois = cells[i + 1].text.strip()

                        # Detect section/type headers like "Bases de données:", etc.
                        if tech.endswith(":") and not mois:
                            current_type = tech.replace(":", "").strip()
                            continue

                        if tech and mois:
                            try:
                                mois_int = int(mois)
                                tech_entries.append({
                                    "technology": tech,
                                    "mois": mois_int,
                                    "type": current_type
                                })
                            except ValueError:
                                pass  # skip invalid mois values

    return tech_entries

# Example usage
results = extract_technologies_with_type("MARTINS_Roni_CV_24-02-13.docx")
