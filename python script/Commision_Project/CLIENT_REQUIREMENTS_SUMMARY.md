# 📋 PROJECT REQUIREMENTS SUMMARY

## Current Understanding of Two Separate Systems:

---

## **SYSTEM 1: HARRY'S GROUP** (Client-Based Fixed Rates)

### How It Works:
- User selects a **CLIENT** from list (Ameristar, Janus, Confidence, Crescent, Medallion HC, Metropolitan)
- Each client has **pre-defined AGENTS** with **FIXED DOLLAR RATES** per plan type
- NO tier system involved - just simple multiplication: `Employee Count × Fixed Rate`

### Example:
```
CLIENT: Confidence
├── Agent 1: Pays $15 per PPC1600/1400/1200 employee, $5 per PPC1000 employee
├── Agent 2: Pays $15 per PPC1600/1400/1200 employee, $5 per PPC1000 employee

If 20 employees on PPC1600:
Agent 1 earns: 20 × $15 = $300
```

### Special Case - Confidence Multipliers:
- Confidence client has different rates based on number of weeks:
  - 2 weeks: $5 (Plan 1000), $15 (Other Plans)
  - 3 weeks: $2.31 (Plan 1000), $5 (Other Plans)
  - 4 weeks: $1.15 (Plan 1000), $3.75 (Other Plans)
  - 5 weeks: $1.15 (Plan 1000), $3 (Other Plans)

### Output Format:
- Charles/Harry/Lighthouse grand totals
- Colored commission columns
- Client-specific agent breakdown

---

## **SYSTEM 2: OTHER GROUPS** (Tier-Based Hierarchy)

### How It Works:
- User provides a **GROUP NAME** (e.g., "100 Academy", "Sales Team Alpha")
- Define a **TOP/MAIN AGENT** with a tier level (e.g., Tier 35)
- Add **SUB-AGENTS** below the main agent, each with their own tier (e.g., Tier 30, 25, 20)
- Uses tier.xlsx rate table for calculations

### Tier Rates Structure:
```
Tier 35: $72 per PPC1600, $53 per PPC1400, $35 per PPC1200, $10 per PPC1000
Tier 30: $52 per PPC1600, $37 per PPC1400, $30 per PPC1200, $8 per PPC1000
Tier 25: $44 per PPC1600, $32 per PPC1400, $25 per PPC1200, $7.5 per PPC1000
Tier 20: $37 per PPC1600, $27 per PPC1400, $20 per PPC1200, $6 per PPC1000
```

### Example Structure:
```
GROUP: 100 Academy
├── Agent 1 (Main/Top): Tier 35
│   └── Manages: 50 employees
│   └── Earns: Tier 35 rates + Override from sub-agents
├── Agent 2 (Sub): Tier 30
│   └── Manages: 30 employees
│   └── Earns: Tier 30 rates only
├── Agent 3 (Sub): Tier 25
│   └── Manages: 20 employees
│   └── Earns: Tier 25 rates only
```

### Override Calculation:
- Main Agent (Tier 35) gets the SPREAD/DIFFERENCE from sub-agents
- If Agent 2 (Tier 30) manages employees:
  - Agent 2 earns: Tier 30 rates
  - Agent 1 gets override: (Tier 35 rates - Tier 30 rates) × Employee Count

### Output Format:
- Simple commission breakdown by agent
- No colored columns
- Agent tier assignments visible
- Main agent override clearly shown

---

## **KEY DIFFERENCES:**

| Feature | Harry's Group | Other Groups |
|---------|---------------|--------------|
| **Selection** | Choose CLIENT name | Enter GROUP name |
| **Rate System** | Fixed dollar amounts | Tier-based rates |
| **Rate Source** | Pre-configured per client | tier.xlsx lookup |
| **Agent Structure** | Pre-defined agents per client | Dynamic: 1 top + multiple sub-agents |
| **Override Logic** | N/A | Main agent earns spread |
| **Output Style** | Charles/Harry/Lighthouse totals, colored | Simple agent list, no colors |

---

## **PROPOSED WORKFLOW:**

### **Option A: Text File Auto-Detection** ✅ (Recommended)
1. User adds payroll Excel files to `Input_Raw/` folder
2. User adds a text file: `Confidence.txt` (just the client name inside)
3. System auto-detects:
   - If `[ClientName].txt` exists → Harry's Group mode
   - If no text file → Prompt for Other Groups configuration
4. Process and generate report

### **Option B: Terminal Menu Selection**
1. User runs script
2. Prompt: "Select mode: [1] Harry's Group [2] Other Groups"
3. If Harry's: "Select client: [1] Ameristar [2] Janus [3] Confidence..."
4. If Other: "Enter group name, top agent tier, sub-agents..."
5. Process and generate report

---

## **QUESTIONS FOR CLIENT CONFIRMATION:**

### 1. ✅ **Is this understanding correct?**
   - Harry's Group = Client selection with fixed rates
   - Other Groups = Tier hierarchy with main + sub-agents

### 2. ✅ **For Harry's Group:**
   - Should we keep the current Charles/Harry/Lighthouse split?
   - Are agent names always "Agent 1", "Agent 2" or custom names?

### 3. ✅ **For Other Groups:**
   - Does the main/top agent ALWAYS earn override from all sub-agents?
   - Can sub-agents have their own sub-agents? (Multi-level) or just 2 levels?

### 4. ✅ **Auto-Detection Method:**
   - Should we use text file method (`Confidence.txt` in Input_Raw folder)?
   - Or terminal menu selection?
   - Or both options available?

### 5. ✅ **Employee Assignment:**
   - For Other Groups: Are ALL employees shared across all agents?
   - Or do specific employees belong to specific agents?

---

**Status:** ⏳ Awaiting Client Confirmation

**Created:** February 2, 2026

**Developer:** Fawaad S.
