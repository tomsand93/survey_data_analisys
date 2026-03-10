"""
Generate sample_data.xlsx — a fully anonymized example file.
Contains 20 fictional respondents with generic placeholder content.
No real survey questions, structure, or client information is included.

Run: python 00_Source_Data/create_sample_data.py
"""
import pandas as pd
import os

# Generic column headers — real column names and positions are not disclosed
HEADERS = [
    'response_id',
    'nps_score',
    'nps_tier',
    'open_q1',           # "What did you like?"
    'open_q1_category',  # LLM-assigned category (empty in raw data)
    'open_q2',           # "How to improve?"
    'open_q2_category',
    'open_q3',           # "One thing to change?"
    'open_q3_category',
    'open_q4',           # "Anything else?"
    'open_q4_category',
    'age_group',
    'gender',
    'segment',
]

respondents = [
    # nps, tier,        q1 (liked),                q2 (improve),               q3 (change),           q4 (else),                  age,    gender,          segment
    (10, 'Promoter',  'Great overall experience',  'More content variety',      'Add new levels',      'Keep it up!',              '25-34', 'Female',        'High Value'),
    ( 9, 'Promoter',  'Very easy to use',          'Faster load times',         'Improved UI',         'Really enjoy it',          '18-24', 'Male',          'Mid Value'),
    ( 8, 'Promoter',  'Social features are fun',   'Minor bug fixes needed',    'Cross-device sync',   'Love the community',       '25-34', 'Male',          'Low Value'),
    ( 7, 'Passive',   'Simple and intuitive',      'Tutorial could be clearer', 'Better onboarding',   'Good but room to grow',    '35-44', 'Female',        'Free'),
    ( 7, 'Passive',   'Rewards keep me engaged',   'Pricing feels high',        'Cheaper options',     'Reduce paywalls please',   '18-24', 'Male',          'Free'),
    ( 6, 'Passive',   'Nice visual design',        'Too many interruptions',    'Ad-free option',      'Ads are disruptive',       '25-34', 'Female',        'Free'),
    ( 5, 'Detractor', 'Interesting concept',       'Crashes on older devices',  'Better optimization', 'Needs stability fixes',    '45-54', 'Male',          'Free'),
    ( 4, 'Detractor', 'Creative content',          'Monetization feels unfair', 'Fairer balance',      'Will not pay more',        '18-24', 'Male',          'Free'),
    ( 9, 'Promoter',  'Fast and responsive',       'More customization',        'New themes',          'Matchmaking is great',     '25-34', 'Female',        'Mid Value'),
    (10, 'Promoter',  'Daily bonuses are great',   'Nothing, love it',          'More events',         'Five stars easily',        '18-24', 'Male',          'High Value'),
    ( 8, 'Promoter',  'Engaging story content',    'More story chapters',       'Story expansion',     'Kept me hooked',           '35-44', 'Female',        'Mid Value'),
    ( 6, 'Passive',   'Solid core experience',     'Moderation improvements',   'Better filters',      'Some bad actors in chat',  '25-34', 'Other',         'Low Value'),
    ( 3, 'Detractor', 'Interesting idea',          'Too many technical issues', 'Fix server errors',   'Uninstalled for now',      '18-24', 'Male',          'Free'),
    ( 9, 'Promoter',  'Well balanced',             'Add replay feature',        'Replay system',       'Competition feels fair',   '25-34', 'Male',          'High Value'),
    ( 7, 'Passive',   'Good variety of features',  'Poor on older hardware',    'Wider device support','Not bad overall',          '35-44', 'Female',        'Free'),
    ( 8, 'Promoter',  'Community events are fun',  'Better group tools',        'Group chat upgrade',  'Events are a highlight',   '18-24', 'Male',          'Mid Value'),
    ( 5, 'Detractor', 'Nice visuals',              'Feels monetization-heavy',  'Fairer pricing',      'Will not spend more',      '25-34', 'Female',        'Free'),
    ( 9, 'Promoter',  'Frequent meaningful updates','Keep doing this',          'More update cadence', 'Team listens to feedback', '18-24', 'Male',          'High Value'),
    ( 6, 'Passive',   'Fun in short sessions',     'Energy limits are strict',  'Remove energy cap',   'Cap is frustrating',       '45-54', 'Prefer not say','Free'),
    ( 8, 'Promoter',  'Smooth controls',           'More accessibility options','Colorblind mode',     'Accessibility matters',    '25-34', 'Female',        'Low Value'),
]

rows = []
for i, (nps, tier, q1, q2, q3, q4, age, gender, segment) in enumerate(respondents, start=1):
    rows.append({
        'response_id':       f'SAMPLE_{i:03d}',
        'nps_score':         nps,
        'nps_tier':          tier,
        'open_q1':           q1,
        'open_q1_category':  '',
        'open_q2':           q2,
        'open_q2_category':  '',
        'open_q3':           q3,
        'open_q3_category':  '',
        'open_q4':           q4,
        'open_q4_category':  '',
        'age_group':         age,
        'gender':            gender,
        'segment':           segment,
    })

df = pd.DataFrame(rows, columns=HEADERS)

out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'sample_data.xlsx')
df.to_excel(out_path, index=False)
print(f"[OK] Sample data written to: {out_path}")
print(f"  {len(respondents)} fictional respondents")
print("  NOTE: All data is fictional and anonymized. No real survey data.")
