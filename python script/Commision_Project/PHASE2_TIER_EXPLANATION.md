# PHASE 2 - TIER-BASED COMMISSION SYSTEM EXPLAINED

## Overview
The tier system is a **hierarchical commission structure** where agents can operate at different tier levels (70, 60, 50, 45, 40, 35, 30, 25, 20, 15). The tier determines:
1. **Base commission** the agent receives
2. **How much of the commission "bubbles up"** to supervisors/uplines at lower tiers
3. The **split** between the agent and their upline


## Understanding the Tier Structure

### The Basic Concept

**Tier 70 = Maximum Commission (100% at your level)**
- Agent at Tier 70 gets the FULL commission for each plan level
- There are no uplines to split with

**Tier 60, 50, 45, etc. = Reduced Commission with Upline Split**
- Agent gets a REDUCED commission
- The "missing" amount goes to an upline agent at a lower tier

### Example: PPC1600 Plans

Looking at the **Tier 70 sheet**:
- **Tier 70 Level**: Agent gets **$107** per PPC1600 plan
  
Looking down the same sheet to other tiers:
- **Tier 60 Level**: Agent gets **$97** + Upline gets **$10** = **$107 total**
- **Tier 50 Level**: Agent gets **$87** + Upline gets **$20** = **$107 total**
- **Tier 45 Level**: Agent gets **$82** + Upline gets **$25** = **$107 total**
- **Tier 40 Level**: Agent gets **$77** + Upline gets **$30** = **$107 total**
- **Tier 35 Level**: Agent gets **$72** + Upline gets **$35** = **$107 total**
- **Tier 30 Level**: Agent gets **$52** + Upline gets **$55** = **$107 total**
- **Tier 25 Level**: Agent gets **$44** + Upline gets **$63** = **$107 total**
- **Tier 20 Level**: Agent gets **$37** + Upline gets **$70** = **$107 total**
- **Tier 15 Level**: Agent gets **$30** + Upline gets **$77** = **$107 total**
- **Tier 10 Level**: Agent gets **$25** + Upline gets **$82** = **$107 total**

---

## Why Multiple Tabs? (The Navigation Path)

Each tab represents a **starting tier level**. The sheet shows:

### **Tier 70 Sheet** (Top Agent)
Shows breakdowns FROM Tier 70 DOWN to Tier 10:
- Tier 70: Full $107
- Tier 60 under 70: $97 to agent, $10 to tier 70 upline
- Tier 50 under 60: $87 to agent, $20 to tier 60 upline
- And so on...

### **Tier 60 Sheet** (Mid-Level Agent)
Shows breakdowns FROM Tier 60 DOWN to Tier 10:
- Tier 60: Full $97 per PPC1600 (their base)
- Tier 50 under 60: $87 to agent, $10 to tier 60 upline
- Tier 45 under 60: $82 to agent, $15 to tier 60 upline
- And so on...

### **Tier 50 Sheet** (Lower Agent)
Shows breakdowns FROM Tier 50 DOWN to Tier 10:
- Tier 50: Full $87 per PPC1600 (their base)
- Tier 45 under 50: $82 to agent, $5 to tier 50 upline
- And so on...

---

## Formula for Commission Calculation

### For an Agent Assigned Tier X:

```
Total Commission = Σ (Number of Employees in Plan × Commission Rate per Plan)
```

Where **Commission Rate** depends on:
1. **Agent's assigned tier** (from dashboard selection)
2. **Employee's plan level** (PPC1600, PPC1400, PPC1200, PPC900/1000)


### Example Calculation

**Scenario:**
- Agent assigned **Tier 60**
- Employee count for this period: 25 PPC1600, 18 PPC1400, 12 PPC1200, 8 PPC900/1000

**Commission = ?**

**Step 1:** Look up rates for Tier 60 from the **Tier 60 sheet** (top section, no upline):
- PPC1600: $97
- PPC1400: $78
- PPC1200: $60
- PPC900/1000: $25

**Step 2:** Multiply employees by rates:
```
Commission = (25 × $97) + (18 × $78) + (12 × $60) + (8 × $25)
           = $2,425 + $1,404 + $720 + $200
           = $4,749
```

---

## Downline Agent Calculation (Subordinate Tier)

If you have an **Agent at Tier 60 with a Downline Agent at Tier 50**:

### Downline Agent Commission:
Look at **Tier 60 sheet → "50 level" section**:
- PPC1600: $87
- PPC1400: $68
- PPC1200: $50
- PPC900/1000: $15

```
Downline Commission = (25 × $87) + (18 × $68) + (12 × $50) + (8 × $15)
                    = $2,175 + $1,224 + $600 + $120
                    = $4,119
```

### Main Agent Bonus (from downline):
```
Main Agent Bonus = Total Commission - Downline Commission
                 = $4,749 - $4,119
                 = $630
```

This $630 is what the Tier 60 agent earns from managing a Tier 50 agent.

---

## Key Differences from Harry's System

| Aspect | Harry's System | Tier System |
|--------|---|---|
| **Rates** | Fixed per client/agent ($5, $10, $15, etc.) | Variable by plan level ($25-$107) |
| **Commission Structure** | Simple: Count × Rate | Hierarchical: Tier-based with upline splits |
| **Upline Management** | Pre-defined clients + agents | Flexible: Multiple agents at different tiers |
| **Tabs** | One tab per date | One tab per tier level (for reference) |
| **Report Output** | Group by client name + agent | Group by agent tier assignment |

---

## Implementation for Phase 2 - Part 2

### Dashboard Features Needed:

1. **Dropdown: Select Group Name**
   - User selects which group (NOT Harry's)

2. **Dynamic Agent Addition**
   - Add Agent 1, Agent 2, Agent 3, etc.
   - Use (+) button to add more agents

3. **Tier Assignment per Agent**
   - Dropdown for each agent: Tier 70, 65, 60, 55, 50, 45, 40, 35, 30, 25, 20, 15
   - Agents can be at DIFFERENT tiers

4. **Commission Calculation**
   - Read employee data from uploaded file
   - Count employees by plan level (PPC1600, 1400, 1200, 900/1000)
   - Look up rate from appropriate tier sheet (60 sheet if agent is Tier 60, etc.)
   - Calculate: (Employee Count × Rate) for each plan level
   - Sum all plans for total commission per agent

5. **Report Output**
   - Single tab showing all agents and their calculated commissions
   - No colored columns (unlike Harry's)
   - Summary of total commissions

---

## Data Structure to Build

```python
TIER_RATES = {
    '70': {
        'PPC1600': 107,
        'PPC1400': 88,
        'PPC1200': 70,
        'PPC900/1000': 25
    },
    '60': {
        'PPC1600': 97,
        'PPC1400': 78,
        'PPC1200': 60,
        'PPC900/1000': 25
    },
    '50': {
        'PPC1600': 87,
        'PPC1400': 68,
        'PPC1200': 50,
        'PPC900/1000': 15
    },
    # ... continue for all tiers
}
```

When Agent is at Tier 60 looking at Tier 50 subordinates:

```python
TIER_DOWNLINE_RATES = {
    '70': {
        'subordinate_60': {'PPC1600': 97, ...},
        'subordinate_50': {'PPC1600': 87, ...},
        # ...
    },
    '60': {
        'subordinate_50': {'PPC1600': 87, ...},
        'subordinate_45': {'PPC1600': 82, ...},
        # ...
    }
}
```

---

## Summary

- **Tier 70**: Full commission, no splits
- **Tier 60**: Partial commission, rest goes up
- **Tier 50, 45, etc.**: Even less, more goes up the chain
- **Why tabs?**: Each tab shows commission breakdown FROM that tier DOWN to lower tiers
- **Calculation**: Count employees × lookup tier rate = commission
- **Dashboard needed**: Group selector → Agent list → Tier assignment → Calculate → Report

---

**Last Updated**: January 30, 2026
**Ready for**: Phase 2 - Part 2 Implementation
