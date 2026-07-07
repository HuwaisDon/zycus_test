from parser.xlsx_adapter import parse_project_xlsx
from parser.rag_engine import analyze_project_health

project = parse_project_xlsx("S2P Project (2).xlsx")
result = analyze_project_health(project)

print("Overall RAG:", result["overall_rag"])
print("Confidence:", result["confidence"])
print("Evidence:")
for e in result["evidence"]:
    print("-", e)

print("\nRecommendations:")
for r in result["recommendations"]:
    print("-", r)

assert result["overall_rag"] in {"Green", "Amber", "Red"}
assert len(result["evidence"]) > 0
assert len(result["recommendations"]) > 0

print("\n✅ Integration test passed!")