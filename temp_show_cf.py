import json
with open('jira_export.json', encoding='utf-8') as f:
    data = json.load(f)
ticket = data[0]
cf = ticket.get('custom_fields', {})
print('=== CUSTOM FIELDS ===')
for k, v in sorted(cf.items()):
    print(f'{k}:')
    print(f'  value: {repr(v.get("value"))[:200]}')
    print(f'  display_name: {repr(v.get("display_name"))}')
