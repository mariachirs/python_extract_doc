def extract_technologies_with_type(doc_path):
    doc = Document(doc_path)
    tech_entries = []
    current_type = None
    table_index = 0

    paragraphs = iter(doc.paragraphs)

    for element in doc.element.body:
        if element.tag.endswith("tbl"):
            table = doc.tables[table_index]
            table_index += 1

            for row in table.rows:
                cells = row.cells
                for i in range(0, len(cells), 2):
                    if i + 1 >= len(cells):
                        continue
                    tech = cells[i].text.strip()
                    mois = cells[i + 1].text.strip()

                    # Detect a type label in a cell (when mois is empty and tech ends in :)
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
                            pass  # Not a valid number

        elif element.tag.endswith("p"):
            paragraph = next(paragraphs)
            text = paragraph.text.strip()
            # Capture section titles like "Bases de données:", etc.
            if text.endswith(":"):
                current_type = text.replace(":", "").strip()

    return tech_entries
