# Email Review Skill

Review recent emails and extract action items organized by person.

## Workflow

1. **Get recent emails** using `get_recent_emails` tool:
   - folder: "Inbox" (or as specified)
   - days: 3 (or as specified)
   - limit: 50

2. **For each email, identify:**
   - Sender name and email
   - Key topic/subject
   - Any explicit requests or asks
   - Deadlines mentioned
   - Questions requiring response
   - Commitments made by others
   - FYIs that need acknowledgment

3. **Categorize action items by type:**
   - **Reply needed** - questions or requests awaiting response
   - **Task to do** - explicit asks with deliverables
   - **Follow up** - items to check back on
   - **Review/Approve** - documents or decisions pending
   - **FYI** - informational, no action required

4. **Output format:**

```markdown
## Email Summary
Reviewed X emails from [folder] (last N days)

## Action Items by Person

### [Person Name] <email@example.com>
- [ ] **Reply needed**: [subject] - [brief description of what's needed]
- [ ] **Task**: [subject] - [what to do] (due: [date if mentioned])

### [Person Name 2]
- [ ] **Follow up**: [subject] - [what to check on]

## No Action Required
- [Person]: [subject] - [one-line summary] (FYI only)

## Flagged/Urgent
- [Any emails marked important or with urgent language]
```

## Example Prompt

"Check my inbox for the last 3 days and give me action items"

## Tool Calls

```
get_recent_emails(folder="Inbox", days=3, limit=50)
```

If user asks for unread only:
```
get_recent_emails(folder="Inbox", days=7, unread_only=true, limit=30)
```

If user asks for flagged:
```
get_recent_emails(folder="Inbox", days=30, flagged_only=true)
```

## Notes

- Group multiple emails from same sender together
- Prioritize emails with questions or explicit requests
- Note if an email thread has multiple back-and-forths
- Flag anything marked high importance
- If email is part of a thread, note if action may already be addressed in later replies
