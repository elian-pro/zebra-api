import json
from zebra_evaluation_builder import generate_evaluation_pdf

with open("ejemplo_evaluacion_cordelia.json", "r", encoding="utf-8") as f:
    eval_data = json.load(f)

output = generate_evaluation_pdf(
    evaluation_data=eval_data,
    output_path="evaluacion_cordelia.pdf",
)

print(f"PDF generado: {output}")
