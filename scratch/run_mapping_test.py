import app
import io

print("--- RUNNING CORE MAPPING ---")
template_file = open("STAND ALONE THEORY EMPTY TEMPLATE.xlsx", "rb")
qp_files = {
    1: open("AAI-CIA1-26-08-2025-set 1.docx", "rb")
}
marks_files = {
    1: open("AML_IA1_Marks.xls", "rb"),
    2: open("AML_IA2_Marks.xls", "rb"),
    3: open("AML_IA3_Marks.xls", "rb")
}
quiz_file = open("AML_IA_Marks (Quiz).xls", "rb")
aat_file = open("AML_IA_Marks(AAT).xls", "rb")

output_bytes = app.process_stand_alone_theory(
    template_file,
    qp_files,
    marks_files,
    quiz_file,
    aat_file
)

with open("test_mapped_out.xlsx", "wb") as f:
    f.write(output_bytes)
print("Mapping complete. Saved to test_mapped_out.xlsx.")
