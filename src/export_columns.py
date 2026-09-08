"""Fixed user-approved source column order; values still come from Jira."""
COLUMNS = [name.strip() for name in '''Summary | Issue key | Issue id | Issue Type | Status | Project key | Project name | Project type | Project lead | Project lead id | Project description | Priority | Resolution | Assignee | Assignee Id | Reporter | Reporter Id | Creator | Creator Id | Created | Updated | Last Viewed | Resolved | Due date | Votes | Description | Environment | Watchers | Watchers Id | Original estimate | Remaining Estimate | Time Spent | Work Ratio | Σ Original Estimate | Σ Remaining Estimate | Σ Time Spent | Security Level | Custom field (Development) | Custom field (Issue colour) | Custom field (Rank) | Sprint | Custom field (Start date) | Custom field (Story point estimate) | Custom field (Team) | Custom field (Vulnerability) | Custom field (Vulnerability) | Status Category | Status Category Changed'''.split('|')]

DIRECT = {'updated': 'Updated', 'lastViewed': 'Last Viewed', 'resolutiondate': 'Resolved',
          'description': 'Description', 'environment': 'Environment', 'workratio': 'Work Ratio',
          'timeoriginalestimate': 'Original estimate', 'timeestimate': 'Remaining Estimate',
          'timespent': 'Time Spent', 'aggregatetimeoriginalestimate': 'Σ Original Estimate',
          'aggregatetimeestimate': 'Σ Remaining Estimate', 'aggregatetimespent': 'Σ Time Spent',
          'security': 'Security Level', 'statuscategorychangedate': 'Status Category Changed',
          'reporter': 'Reporter', 'creator': 'Creator', 'votes': 'Votes', 'watches': 'Watchers'}


def fixed_headers(mapping, existing):
    result, used = [], set()
    for label in COLUMNS:
        if label in used:
            alternatives = [h for h in mapping.values() if h.startswith(label + ' [') and h not in used]
            label = alternatives[0] if alternatives else label + ' [2]'
        result.append(label)
        used.add(label)
    return result + [h for h in existing if h not in used]
