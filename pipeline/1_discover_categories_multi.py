import pandas as pd
import requests
import json
import sys
import io
import os
from collections import defaultdict

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

print("="*80)
print("VERSION 6: MULTI-CATEGORY LLM DISCOVERY")
print("="*80)
print("Key Feature: Each response can have MULTIPLE categories")
print("Example: 'Graphics and prizes' → 'Graphics | Rewards & Boosters'")
print("="*80)

# Configuration
OLLAMA_URL = "http://localhost:11434/api/generate"
OLLAMA_MODEL = "llama3.2"
SAMPLE_SIZE = 200  # Sample 200 responses per question to discover categories

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
output_file = os.path.join(base_dir, 'intermediate', 'discovered_categories_multi.json')

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
        return True

    text = str(response_text).strip().lower()
    text_clean = text.replace('.', '').replace(',', '').replace('!', '')

    trivial = [
        'no', 'nope', 'na', 'n/a', 'nothing', 'none', 'not really',
        'nah', 'nada', 'idk', 'dont know', 'not sure', 'no idea'
    ]

    return text_clean in trivial or len(text_clean) <= 2


def discover_categories_with_llm(question, response_text):
    """
    Use LLM to discover ONE OR MORE categories for a response.
    Returns list of categories.
    """
    prompt = f"""You are analyzing survey responses about a mobile game.

Question: {question}
Response: "{response_text}"

Identify ONE OR MORE themes/categories that this response mentions. A response can cover multiple topics.

Rules:
1. List ALL distinct themes mentioned (e.g., if they mention graphics AND rewards, list both)
2. Each category should be 2-5 words
3. Be specific but not overly narrow
4. Output ONLY the category names, one per line
5. If multiple aspects are mentioned, list each separately

Examples:
- "I love the graphics and the daily rewards" → Graphics & Visuals
Rewards & Progression
- "More boosters and better matchmaking" → Booster System
Matchmaking & Fairness
- "The game is fun" → Overall Enjoyment

Categories (one per line):"""

    try:
        response = requests.post(
            OLLAMA_URL,
            json={
                "model": OLLAMA_MODEL,
                "prompt": prompt,
                "stream": False,
                "options": {
                    "temperature": 0.4,
                    "num_predict": 100
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
                if line and len(line) > 2 and len(line) < 60:
                    categories.append(line)

            # Limit to top 3 categories max
            return categories[:3] if categories else ["Other"]
        else:
            return ["Other"]

    except Exception as e:
        print(f"\n⚠️  Error calling Ollama: {e}")
        return ["Other"]


# Process each question
print("\n" + "="*80)
print("DISCOVERING CATEGORIES WITH MULTI-CATEGORY SUPPORT")
print("="*80)

discovered_categories = {}

for q_id, q_info in open_questions.items():
    print(f"\n{'='*80}")
    print(f"Processing {q_id}: {q_info['question']}")
    print(f"{'='*80}")

    answer_col = q_info['answer_col']

    # Get substantive responses (filter out trivial)
    substantive_mask = responses.iloc[:, answer_col].notna()
    substantive_responses = responses[substantive_mask]

    # Filter out trivial responses
    substantive_responses = substantive_responses[
        ~substantive_responses.iloc[:, answer_col].apply(is_trivial_response)
    ]

    # Sample
    if len(substantive_responses) > SAMPLE_SIZE:
        sample = substantive_responses.sample(n=SAMPLE_SIZE, random_state=42)
    else:
        sample = substantive_responses

    print(f"\nTotal responses: {len(responses)}")
    print(f"Substantive responses: {len(substantive_responses)}")
    print(f"Sample size: {len(sample)}")

    # Discover categories
    print(f"\nDiscovering categories (this will take a while)...")

    category_count = defaultdict(int)

    for idx, (row_idx, row) in enumerate(sample.iterrows(), 1):
        response_text = row.iloc[answer_col]

        # Get multiple categories from LLM
        categories = discover_categories_with_llm(q_info['question'], response_text)

        # Count each category
        for category in categories:
            category_count[category] += 1

        # Progress indicator
        if idx % 10 == 0 or idx == len(sample):
            print(f"  Progress: {idx}/{len(sample)} ({idx/len(sample)*100:.1f}%)", end='\r')

    print()  # New line

    # Sort categories by frequency
    sorted_categories = sorted(category_count.items(), key=lambda x: x[1], reverse=True)

    print(f"\n✓ Discovered {len(sorted_categories)} categories")
    print(f"\nTop 20 categories:")
    for i, (cat, count) in enumerate(sorted_categories[:20], 1):
        print(f"  {i}. {cat}: {count} mentions")

    # Store results
    discovered_categories[q_id] = {
        'question': q_info['question'],
        'categories': [cat for cat, count in sorted_categories],
        'category_counts': {cat: count for cat, count in sorted_categories}
    }

# Save results
print(f"\n{'='*80}")
print("SAVING DISCOVERED CATEGORIES")
print(f"{'='*80}")

with open(output_file, 'w', encoding='utf-8') as f:
    json.dump(discovered_categories, f, indent=2, ensure_ascii=False)

print(f"\n✓ Saved to: {output_file}")
print(f"\nSummary:")
for q_id, data in discovered_categories.items():
    print(f"  {q_id}: {len(data['categories'])} categories discovered")

print(f"\n{'='*80}")
print("DISCOVERY COMPLETE")
print(f"{'='*80}")
print(f"\nNext step:")
print(f"  Run 2_refine_categories.py to review and refine the categories")
print(f"\nKey Feature:")
print(f"  Each response can now be assigned MULTIPLE categories")
print(f"  This captures the full richness of feedback!")
