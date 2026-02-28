from docx import Document
import sys
import os


def generate_rbt(minggu, kelas, tarikh, hari, masa, refleksi):

    # ==============================
    # 1️⃣ Load template ikut minggu
    # ==============================
    template_path = f"rbt_templates/M{minggu}.docx"

    if not os.path.exists(template_path):
        raise Exception("Template minggu tidak dijumpai.")

    doc = Document(template_path)

    # ==============================
    # 2️⃣ Replace placeholder
    # ==============================
    for p in doc.paragraphs:
        if "{{KELAS}}" in p.text:
            p.text = p.text.replace("{{KELAS}}", kelas)

        if "{{TARIKH}}" in p.text:
            p.text = p.text.replace("{{TARIKH}}", tarikh)

        if "{{HARI}}" in p.text:
            p.text = p.text.replace("{{HARI}}", hari)

        if "{{MASA}}" in p.text:
            p.text = p.text.replace("{{MASA}}", masa)

        if "{{REFLEKSI}}" in p.text:
            p.text = p.text.replace("{{REFLEKSI}}", refleksi)

    # ==============================
    # 3️⃣ Replace dalam TABLE juga
    # ==============================
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "{{KELAS}}" in cell.text:
                    cell.text = cell.text.replace("{{KELAS}}", kelas)

                if "{{TARIKH}}" in cell.text:
                    cell.text = cell.text.replace("{{TARIKH}}", tarikh)

                if "{{HARI}}" in cell.text:
                    cell.text = cell.text.replace("{{HARI}}", hari)

                if "{{MASA}}" in cell.text:
                    cell.text = cell.text.replace("{{MASA}}", masa)

                if "{{REFLEKSI}}" in cell.text:
                    cell.text = cell.text.replace("{{REFLEKSI}}", refleksi)

    # ==============================
    # 4️⃣ Save output
    # ==============================
    output_folder = "output"
    os.makedirs(output_folder, exist_ok=True)

    output_file = f"{output_folder}/RPH_RBT_M{minggu}.docx"
    doc.save(output_file)

    return output_file


# ==================================
# RUN DARI TERMINAL / NODE
# ==================================
if __name__ == "__main__":

    minggu = sys.argv[1]
    kelas = sys.argv[2]
    tarikh = sys.argv[3]
    hari = sys.argv[4]
    masa = sys.argv[5]
    refleksi = sys.argv[6]

    file_path = generate_rbt(minggu, kelas, tarikh, hari, masa, refleksi)

    print("SIAP:", file_path)
