import json
import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

print("="*80)
print("CATEGORY REFINEMENT TOOL")
print("="*80)
print("Purpose: Review and refine LLM-discovered categories")
print("Method: Accept as-is or manually edit discovered_categories.json")
print("="*80)

# Load discovered categories
script_dir = os.path.dirname(os.path.abspath(__file__))
base_dir = os.path.dirname(script_dir)  # Go up two levels from 
input_file = os.path.join(base_dir, 'intermediate', 'discovered_categories_multi.json')
output_file = os.path.join(base_dir, 'intermediate', 'refined_categories_multi.json')

if not os.path.exists(input_file):
    print(f"\n❌ ERROR: {input_file} not found!")
    print("   Please run discover_categories_llm.py first")
    sys.exit(1)

with open(input_file, 'r', encoding='utf-8') as f:
    discovered = json.load(f)

print(f"\nLoaded discovered categories from: {input_file}")

refined_categories = {}

for q_id, data in discovered.items():
    print(f"\n{'='*80}")
    print(f"{q_id}: {data['question']}")
    print(f"{'='*80}")

    categories = data['categories']  # This is a list
    category_counts = data.get('category_counts', {})  # This is a dict with counts
    print(f"\nDiscovered {len(categories)} categories:")

    # Sort by count (descending) if counts available, otherwise just use order
    if category_counts:
        sorted_cats = sorted(category_counts.items(), key=lambda x: -x[1])
    else:
        sorted_cats = [(cat, 0) for cat in categories]

    for i, (cat, count) in enumerate(sorted_cats, 1):
        print(f"  {i:2d}. {cat:45s} ({count:3d} mentions)")

print(f"\n{'='*80}")
print("REFINEMENT OPTIONS")
print("="*80)
print("\nYou have two options:")
print("\n  1. ACCEPT AS-IS")
print("     Keep all discovered categories without changes")
print("     (Script will automatically create refined_categories.json)")
print("\n  2. MANUAL REFINEMENT")
print("     Edit discovered_categories.json to:")
print("     - Merge similar categories (combine counts)")
print("     - Rename categories for clarity")
print("     - Remove unwanted categories")
print("     Then re-run this script")

choice = input(f"\nChoice [1=Accept, 2=Manual edit]: ").strip()

if choice == '1':
    print("\n✓ Accepting all categories as-is")

    for q_id, data in discovered.items():
        refined_categories[q_id] = {
            'question': data['question'],
            'categories': data['categories']  # Already a list, no need to call .keys()
        }

    # Save refined categories
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(refined_categories, f, indent=2, ensure_ascii=False)

    print(f"\n{'='*80}")
    print("REFINEMENT COMPLETE")
    print("="*80)
    print(f"✓ Refined categories saved to: {output_file}")
    print(f"\nNext step:")
    print(f"  Run: python pipeline/3_categorize_multi.py")
    print("="*80)

elif choice == '2':
    print(f"\n{'='*80}")
    print("MANUAL REFINEMENT INSTRUCTIONS")
    print("="*80)
    print(f"\n1. Open: {input_file}")
    print(f"\n2. Edit the categories as needed:")
    print(f"   - To merge: Change category names to be identical, sum counts")
    print(f"   - To rename: Just change the category name")
    print(f"   - To remove: Delete the category entry")
    print(f"\n3. Save the file")
    print(f"\n4. Re-run this script: python pipeline/2_refine_categories.py")
    print(f"\n{'='*80}")

else:
    print("\n❌ Invalid choice. Please run again and enter 1 or 2")
