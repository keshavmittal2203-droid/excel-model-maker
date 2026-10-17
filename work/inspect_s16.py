import json

pack = json.load(open("proposals/err4.deck-pack.proposal.json"))
inputs = json.load(open("proposals/err4.inputs.proposal.json"))

s16_slots = [s for s in pack["slots"] if s["slide"] == 16]
print("Slide 16 slots count:", len(s16_slots))
for s in s16_slots:
    name = s["name"]
    val = inputs["values"].get(name, {}).get("value")
    src = inputs["values"].get(name, {}).get("source")
    out_str = f"{name} ({s.get('target_type')}): {val} [source: {src}]"
    print(out_str.encode("ascii", "backslashreplace").decode("ascii"))
