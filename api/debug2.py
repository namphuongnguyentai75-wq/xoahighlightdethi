import docx
import re
import glob

files = glob.glob('/tmp/tmp*/Y17*')
input_file = sorted(files)[-1]
print(f"Analyzing: {input_file}")

doc = docx.Document(input_file)

# Simulate paragraph deletion logic and show what would be deleted
deleted = []
kept = []

for p_element in doc.element.xpath('.//w:p'):
    texts = p_element.xpath('.//w:t')
    full_text = "".join(t.text for t in texts if t.text)
    full_text_lower = full_text.lower().strip()

    if not full_text_lower:
        continue

    # Check protection
    is_protected = False
    if re.match(r'^\d+\.?\s', full_text.strip()):
        is_protected = True
    if re.match(r'^[A-Da-d][\.\)]\s', full_text.strip()):
        is_protected = True

    # Check deletion patterns
    should_delete = False
    reason = ""
    if re.search(r'item\s*\d+\s*:\s*score', full_text_lower):
        should_delete = True
        reason = "item score"
    elif re.match(r'^difficulty', full_text_lower):
        should_delete = True
        reason = "difficulty"
    elif re.match(r'^proportion\b', full_text_lower):
        should_delete = True
        reason = "proportion"
    elif re.match(r'^highest\s+group\b', full_text_lower):
        should_delete = True
        reason = "highest group"
    elif re.match(r'^lowest\s+group\b', full_text_lower):
        should_delete = True
        reason = "lowest group"
    elif re.match(r'^r-pbis\b', full_text_lower):
        should_delete = True
        reason = "r-pbis"
    elif re.match(r'^p\s+value\b', full_text_lower):
        should_delete = True
        reason = "p value"

    if should_delete and not is_protected:
        deleted.append((reason, full_text[:100]))
    elif should_delete and is_protected:
        print(f"PROTECTED from deletion: [{reason}] {full_text[:100]}")

print(f"\nTotal paragraphs that would be deleted: {len(deleted)}")
print(f"\n=== Deleted paragraphs (first 30): ===")
for reason, text in deleted[:30]:
    print(f"  [{reason}] {text}")

print(f"\n=== Deleted paragraphs (last 10): ===")
for reason, text in deleted[-10:]:
    print(f"  [{reason}] {text}")

# Count by reason
from collections import Counter
reasons = Counter(r for r, _ in deleted)
print(f"\n=== Deletion reasons: ===")
for r, c in reasons.most_common():
    print(f"  {r}: {c}")
