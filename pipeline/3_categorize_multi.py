import pandas as pd
import requests
import json
import sys
import io
import time
import re
import os
from collections import defaultdict

# Store example quotes per question & category
example_quotes = defaultdict(dict)

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

print("="*80)
print("VERSION 6: MULTI-CATEGORY CATEGORIZATION")
print("="*80)
print("Key Feature: Assigns MULTIPLE categories per response")
print("Storage: Categories separated by ' | ' delimiter")
print("="*80)

# Configuration
OLLAMA_URL = "http://localhost:11434/api/generate"
OLLAMA_MODEL = "llama3.2"

# Test Ollama connection
print("\nTesting Ollama connection...")
try:
    test_response = requests.post(
        OLLAMA_URL,
        json={"model": OLLAMA_MODEL, "prompt": "Say 'OK'", "stream": False},
        timeout=10
    )
    if test_response.status_code == 200:
        print(f"✓ Ollama is running and {OLLAMA_MODEL} model is available")
    else:
        print(f"⚠️  Ollama returned status code: {test_response.status_code}")
        sys.exit(1)
except requests.exceptions.RequestException as e:
    print(f"❌ Cannot connect to Ollama at {OLLAMA_URL}")
    print(f"   Error: {e}")
    sys.exit(1)

# Set up paths
script_dir = os.path.dirname(os.path.abspath(__file__))
base_dir = os.path.dirname(script_dir)
data_file = os.path.join(base_dir, 'data', 'full_data.xlsx')
categories_file = os.path.join(base_dir, 'intermediate', 'refined_categories_multi.json')
output_file = os.path.join(base_dir, 'processed_data', 'full_data_multi_category.xlsx')

# Load refined categories
if not os.path.exists(categories_file):
    print(f"\n❌ ERROR: {categories_file} not found!")
    print("   Please run 1_discover_categories_multi.py and 2_refine_categories.py first")
    sys.exit(1)

with open(categories_file, 'r', encoding='utf-8') as f:
    refined_categories = json.load(f)

print(f"✓ Loaded refined categories from: {categories_file}")

# Load survey data
print(f"\nLoading survey data from: {data_file}")
df = pd.read_excel(data_file)
metadata_row = df.iloc[0]
responses = df.iloc[1:].copy()

print(f"✓ Loaded data: {len(responses)} respondents, {len(df.columns)} columns")

# Define open-ended questions
open_questions = {
    'Q3': {
        'question': 'What did you like most about your experience?',
        'answer_col': 5,
        'insight_col': 6
    },
    'Q4': {
        'question': 'How can we improve your experience?',
        'answer_col': 7,
        'insight_col': 8
    },
    'Q12': {
        'question': 'If exactly ONE thing could be changed or added to the game what would it be?',
        'answer_col': 41,
        'insight_col': 42
    },
    'Q15': {
        'question': 'Anything else you\'d like to share before we wrap up?',
        'answer_col': 48,
        'insight_col': 49
    }
}


def is_trivial_response(response_text):
    """Check if response is trivial (no/nope/na/nothing)"""
    if pd.isna(response_text):
        return False

    text = str(response_text).strip().lower()
    text_clean = re.sub(r'[^\w\s]', '', text)

    trivial = [
        'no', 'nope', 'na', 'n/a', 'nothing', 'none', 'not really',
        'nah', 'nada', 'idk', 'dont know', 'not sure', 'no idea', ''
    ]

    return text_clean in trivial or len(text_clean) <= 2


def categorize_with_llm_multi(question, response_text, existing_categories, q_id):
    """
    Use Ollama LLM to categorize a response into ONE OR MORE categories.
    Returns list of categories.
    """
    # Build category list
    category_list = "\n".join([f"- {cat}" for cat in existing_categories[:25]])

    prompt = f"""You are categorizing survey responses. Assign ONE OR MORE categories to this response.

Question: {question}
Response: "{response_text}"

Existing categories:
{category_list}

Instructions:
1. Assign ALL relevant categories from the list above (a response can have multiple themes)
2. Only use categories from the list above
3. If none fit perfectly, choose the closest match(es)
4. Output ONLY category names, one per line
5. Maximum 3 categories per response

Examples:
Response: "Love the graphics and rewards"
Categories:
Graphics & Visuals
Rewards & Progression

Response: "More boosters and better matchmaking please"
Categories:
Booster System
Matchmaking & Fairness

Categories (one per line):"""

    try:
        response = requests.post(
            OLLAMA_URL,
            json={
                "model": OLLAMA_MODEL,
                "prompt": prompt,
                "stream": False,
                "options": {
                    "temperature": 0.3,
                    "num_predict": 80
                }
            },
            timeout=45
        )

        if response.status_code == 200:
            result = response.json()
            llm_output = result.get('response', '').strip()

            # Parse output - split by newlines and clean
            categories = []
            for line in llm_output.split('\n'):
                line = line.strip()
                # Remove bullet points, numbers, etc.
                line = line.lstrip('•-*123456789. ')

                # Check if this matches an existing category
                for existing_cat in existing_categories:
                    if line.lower() == existing_cat.lower() or existing_cat.lower() in line.lower():
                        if existing_cat not in categories:  # Avoid duplicates
                            categories.append(existing_cat)
                        break

            # Limit to top 3 categories
            if len(categories) > 3:
                categories = categories[:3]

            return categories if categories else ["Other"]
        else:
            return ["Other"]

    except Exception as e:
        print(f"\n⚠️  Error calling Ollama: {e}")
        return ["Other"]


# Process each open-ended question
print("\n" + "="*80)
print("CATEGORIZING WITH MULTI-CATEGORY SUPPORT")
print("="*80)

total_categorized = 0
total_trivial = 0
total_llm = 0
total_multi_category = 0  # Count responses with 2+ categories

for q_id, q_info in open_questions.items():
    print(f"\n{'='*80}")
    print(f"Processing {q_id}: {q_info['question']}")
    print(f"{'='*80}")

    # Get categories for this question
    if q_id not in refined_categories:
        print(f"⚠️  No categories found for {q_id}, skipping")
        continue

    existing_categories = refined_categories[q_id]['categories']
    print(f"Available categories: {len(existing_categories)}")
    print(f"Top 5: {existing_categories[:5]}")

    # Extract response and insight columns
    answer_col_idx = q_info['answer_col']
    insight_col_idx = q_info['insight_col']

    # Find all responses (categorize everyone)
    all_responses_mask = pd.Series([True] * len(responses), index=responses.index)
    all_indices = responses[all_responses_mask].index

    print(f"\nTotal responses to categorize: {len(all_indices)}")

    # Categorize each response
    q_trivial = 0
    q_llm = 0
    q_multi = 0

    for idx, row_idx in enumerate(all_indices, 1):
        response_text = responses.at[row_idx, responses.columns[answer_col_idx]]

        # Check if no answer provided
        if pd.isna(response_text):
            categories = ["No Response"]
        # Check if trivial response
        elif is_trivial_response(response_text):
            categories = ["No Additional Feedback"]
            q_trivial += 1
        else:
            categories = categorize_with_llm_multi(
                q_info['question'],
                response_text,
                existing_categories,
                q_id
            )
            q_llm += 1

            if len(categories) > 1:
                q_multi += 1

    # Save example quote per category (first occurrence only)
    for cat in categories:
        if cat not in example_quotes[q_id]:
            example_quotes[q_id][cat] = str(response_text).strip()


        # Join categories with delimiter
        category_string = " | ".join(categories)

        # Assign to insight column
        responses.at[row_idx, responses.columns[insight_col_idx]] = category_string

        # Progress indicator
        if idx % 10 == 0 or idx == len(all_indices):
            print(f"  Progress: {idx}/{len(all_indices)} ({idx/len(all_indices)*100:.1f}%) "
                  f"[Trivial: {q_trivial}, LLM: {q_llm}, Multi: {q_multi}]", end='\r')

        # Small delay to avoid overwhelming Ollama
        if q_llm > 0 and q_llm % 5 == 0:
            time.sleep(0.1)

    print()  # New line after progress
    print(f"✓ Categorized {len(all_indices)} responses")
    print(f"  - Trivial responses: {q_trivial}")
    print(f"  - LLM categorized: {q_llm}")
    print(f"  - Multi-category responses: {q_multi} ({q_multi/q_llm*100:.1f}% of substantive)")

    total_categorized += len(all_indices)
    total_trivial += q_trivial
    total_llm += q_llm
    total_multi_category += q_multi

# Reconstruct the full dataframe with metadata row
df_enhanced = pd.concat([metadata_row.to_frame().T, responses], ignore_index=True)

# Save enhanced data
print(f"\n{'='*80}")
print("SAVING MULTI-CATEGORY DATA")
print(f"{'='*80}")

df_enhanced.to_excel(output_file, index=False, engine='openpyxl')


# ============================================================
# SAVE EXAMPLE QUOTES PER CATEGORY
# ============================================================

examples_output = os.path.join(
    base_dir,
    'output',
    'category_example_quotes.xlsx'
)

os.makedirs(os.path.dirname(examples_output), exist_ok=True)

with pd.ExcelWriter(examples_output, engine='openpyxl') as writer:
    for q_id, cat_map in example_quotes.items():
        rows = []
        for cat, quote in cat_map.items():
            rows.append({
                "Category": cat,
                "Example Quote": quote
            })

        if rows:
            df_examples = pd.DataFrame(rows)
            df_examples.to_excel(
                writer,
                sheet_name=q_id,
                index=False
            )

print(f"\n✓ Saved category example quotes to: {examples_output}")


print(f"\n✓ Saved multi-category data to: {output_file}")
print(f"\nSummary:")
print(f"  Total responses categorized: {total_categorized}")
print(f"  - Trivial responses: {total_trivial}")
print(f"  - LLM categorized: {total_llm}")
print(f"  - Multi-category responses: {total_multi_category}")
print(f"  - Multi-category rate: {total_multi_category/total_llm*100:.1f}% of substantive")

print(f"\n{'='*80}")
print("MULTI-CATEGORY CATEGORIZATION COMPLETE")
print(f"{'='*80}")
print(f"\nKey Achievement:")
print(f"  {total_multi_category} responses captured with multiple categories!")
print(f"  Example: 'Graphics & Visuals | Rewards & Boosters'")
print(f"\nNext step:")
print(f"  Run 4_analyze_multi_categories.py to generate analysis")
