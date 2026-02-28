# Construction Planning & Scheduling Knowledge Base

## 1. Schedule Health Assessment Framework

### DCMA 14-Point Assessment — Detailed Implementation

#### Check 1: Logic (Open Ends)
**Metric**: % of incomplete tasks missing predecessor OR successor
**Threshold**: < 5%
**Implementation**:
```
For each task where percent_complete < 100:
  - Must have >= 1 predecessor (except project start milestone)
  - Must have >= 1 successor (except project finish milestone)
  - Summary tasks: logic on children counts, not on summary itself
```
**Common Issues**: Dangling activities from imported schedules, deleted tasks leaving orphans

#### Check 2: Leads (Negative Lag)
**Metric**: % of relationships with negative lag
**Threshold**: 0% (no leads allowed)
**Why**: Leads create artificial overlap that hides true logic. Replace with SS + positive lag.
**Fix**: Convert `FS-5d` to `SS+offset` relationship

#### Check 3: Lags (Positive Lag)
**Metric**: % of relationships with positive lag
**Threshold**: < 5%
**Why**: Lags often mask missing activities (cure time, approval time, procurement)
**Fix**: Create explicit activities for wait periods (e.g., "Concrete Curing - 7 days")

#### Check 4: Relationship Types
**Metric**: % of non-FS relationships
**Threshold**: < 10%
**Why**: FS is most intuitive and auditable. Excessive SS/FF/SF indicates poor planning.
**Note**: SS links for parallel work zones are acceptable (e.g., floor-by-floor progression)

#### Check 5: Hard Constraints
**Metric**: % of incomplete tasks with Must Start On / Must Finish On constraints
**Threshold**: < 5%
**Why**: Hard constraints override CPM logic, hide true float, prevent accurate forecasting
**Acceptable**: Contractual milestones, regulatory deadlines, third-party dependencies
**Fix**: Use SNET/SNLT (soft constraints) where possible

#### Check 6: High Float
**Metric**: % of incomplete tasks with Total Float > 44 working days
**Threshold**: < 5%
**Why**: Excessive float indicates missing logic or poorly connected activities
**Common Causes**: Activities connected to project finish but not to preceding work

#### Check 7: Negative Float
**Metric**: Number of tasks with negative Total Float
**Threshold**: 0 tasks
**Why**: Negative float = schedule is late. Must address immediately.
**Fix**: Crashing, fast-tracking, re-sequencing, scope reduction

#### Check 8: High Duration
**Metric**: % of incomplete tasks with duration > 44 working days
**Threshold**: < 5%
**Why**: Long activities reduce visibility and control
**Fix**: Break into smaller work packages (2-4 week activities)

#### Check 9: Invalid Dates
**Metric**: Tasks with actual start/finish dates in the future (after report date)
**Threshold**: 0 tasks
**Why**: Data integrity — actuals cannot be in the future
**Common Cause**: Incorrect progress entry, bad date formatting

#### Check 10: Resources
**Metric**: % of incomplete tasks without resource assignments
**Threshold**: < 5% (for resource-loaded schedules)
**Why**: Resource loading enables: leveling, S-curves, earned value, histograms
**Note**: Summary tasks and milestones exempt from this check

#### Check 11: Missed Tasks
**Metric**: Incomplete tasks with forecast finish date < report date (should be complete)
**Threshold**: 0 tasks
**Why**: Indicates progress not being reported or schedule not updated
**Fix**: Update actual dates or re-forecast

#### Check 12: Critical Path Test
**Metric**: Valid, continuous, and reasonable critical path
**Threshold**: PASS
**Checks**:
- CP exists and runs from project start to finish
- CP is continuous (no breaks in logic)
- CP length is reasonable (typically 60-80% of total project duration)
- CP has no hard-constrained activities breaking the chain

#### Check 13: CPLI (Critical Path Length Index)
**Formula**: CPLI = (Critical Path Length + Total Float) / Critical Path Length
**Threshold**: >= 1.0
**Interpretation**:
- CPLI >= 1.0: Schedule can meet deadline
- CPLI < 1.0: Schedule is behind, cannot meet deadline without compression
- The closer to 1.0, the tighter the schedule

#### Check 14: BEI (Baseline Execution Index)
**Formula**: BEI = # Tasks completed on time / # Tasks planned to be complete
**Threshold**: >= 0.95
**Interpretation**:
- BEI = 1.0: Executing exactly as planned
- BEI < 0.95: Falling behind plan — investigate root causes

---

## 2. Forensic Schedule Analysis Methods

### Method 1: As-Planned vs As-Built
- Compare original baseline schedule with actual completion
- Simple but doesn't show cause-and-effect
- Used for: small claims, quick overview

### Method 2: Impacted As-Planned
- Add delay events to as-planned schedule
- Forward pass shows theoretical impact
- Prospective method — used for: delay claims, time extensions
- Weakness: Doesn't account for mitigation or actual progress

### Method 3: Collapsed As-Built (But-For Analysis)
- Remove delay events from as-built schedule
- Backward calculation shows "but for" completion date
- Retrospective method — used for: complex delay claims
- Strength: Based on actual events

### Method 4: Time Impact Analysis (TIA)
- Preferred by AACE (52R-06) and many contracts
- Step-by-step insertion of delay events at their actual occurrence
- Shows cumulative impact on critical path
- Requirements: Updated schedule at each delay event, contemporaneous records

### Method 5: Windows Analysis
- Divide project into time windows (typically monthly)
- Analyze critical path and delays within each window
- Most thorough method for concurrent delay analysis
- Used for: complex multi-party delay claims

### Method 6: Longest Path Analysis
- Identify the longest path through the network (may differ from CPM critical path)
- Accounts for actual durations vs planned
- Useful when schedule has multiple near-critical paths

---

## 3. Resource Management Best Practices

### Resource Types in Asta
- **Permanent Resources**: Labour (carpenters, electricians, crane operators) and equipment (cranes, excavators)
- **Consumable Resources**: Materials with quantity tracking (concrete m3, steel tonnes)
- **Cost Centres**: Budget categories (preliminaries, substructure, superstructure)

### Resource Leveling Strategies
1. **Unlimited Resources**: CPM-only, no resource constraints
2. **Resource Constrained**: Delay activities to stay within resource limits
3. **Resource Smoothed**: Minimize peaks while respecting critical path
4. **Time-Limited**: Fixed end date, add resources to resolve overallocation

### S-Curve Analysis
- **Planned S-Curve**: Cumulative planned progress over time
- **Actual S-Curve**: Cumulative actual progress
- **Earned Value S-Curve**: BCWP/BCWS comparison
- **Resource Histogram**: Period-by-period resource demand
- Early start and late start envelopes show the "banana curve" (acceptable range)

---

## 4. Progress Tracking Methods

### Duration-Based Progress
- Simple: % of duration elapsed
- Suitable for activities with uniform work distribution
- `percent_complete = actual_duration / planned_duration * 100`

### Physical Progress (Quantities)
- Based on measurable output (m2 of formwork, m3 of concrete, units installed)
- Most accurate for construction activities
- `percent_complete = actual_quantity / planned_quantity * 100`

### Milestone Weighting
- Assign weight to each milestone within an activity
- Progress = sum of completed milestone weights
- Example: Rebar (20%) + Formwork (20%) + Pour (30%) + Strip (15%) + Cure (15%) = 100%

### Earned Value Progress
- BCWP-based progress
- Integrates cost and schedule performance
- Best for project-level reporting

---

## 5. Programme Submission Requirements

### NEC4 Requirements (Clause 31)
A programme must show:
- Activities with planned start/finish dates
- Order and timing of operations
- Float and time risk allowances (TRA)
- Health and safety requirements
- Procedures for design, procurement, testing
- Dates of key milestones and constraints
- Resource histograms
- Method statements (where relevant)
- Critical path clearly identified

### FIDIC Requirements (Clause 8.3)
A programme must show:
- Order in which contractor intends to carry out works
- Anticipated timing of each stage
- Critical path
- Resource histograms (labour and equipment)
- Quality assurance procedures
- Cash flow forecast

### Typical Client/Contract Requirements
- Monthly programme updates within 7-14 days of report date
- Baseline comparison on all submissions
- Narrative report explaining: progress, delays, mitigation, forecast
- 3-week / 6-week look-ahead programmes
- Resource histograms and S-curves
- Risk-adjusted programme (P50/P80 scenarios)

---

## 6. Construction-Specific Scheduling Techniques

### Line of Balance (LOB) / Repetitive Scheduling
- For projects with repeating units (residential floors, hotel rooms, pipeline sections)
- Each trade progresses through units at a planned rate
- Identifies: bottleneck trades, buffer requirements, optimal crew sizes
- Asta Powerproject supports LOB natively via "chart types"

### 4D BIM Integration
- Link schedule activities to 3D model elements
- Visualize construction sequence
- Clash detection: spatial conflicts between trades at same time/location
- Asta supports IFC model import and 4D simulation

### Last Planner System (LPS)
- Collaborative planning methodology
- Phase planning → Look-ahead planning → Weekly work planning → Daily coordination
- PPC (Percent Plan Complete) = completed tasks / committed tasks
- Reasons for non-completion analysis
- Not directly in Asta but schedule structure should support it

### Pull Planning
- Start from completion milestone, work backwards
- Each trade defines what they need from predecessors
- Results in demand-driven schedule with realistic handoffs
- Particularly effective for MEP-heavy projects

---

## 7. Key Construction Metrics

### Schedule Metrics
| Metric | Formula | Target |
|--------|---------|--------|
| SPI | EV / PV | >= 1.0 |
| CPLI | (CP + TF) / CP | >= 1.0 |
| BEI | Completed / Planned | >= 0.95 |
| PPC | Completed / Committed | >= 80% |
| Float Consumption Rate | Float used / Time elapsed | < 1.0 |

### Production Rates (Typical UK/EU)
| Activity | Typical Rate | Unit |
|----------|-------------|------|
| Piling (CFA) | 6-12 piles/day | nr/day |
| Concrete pour | 40-80 m3/day | m3/day |
| RC frame cycle | 5-7 days/floor | days/floor |
| Steel erection | 20-40 tonnes/day | t/day |
| Brickwork | 40-60 m2/day (per gang) | m2/day |
| Curtain walling | 15-25 m2/day | m2/day |
| Drylining | 20-30 m2/day (per gang) | m2/day |
| MEP first fix | 1-2 weeks/floor | weeks/floor |
| Floor screeds | 200-400 m2/day | m2/day |

---

## 8. Weather and Seasonal Considerations

### Temperature Constraints
- Concrete: cannot pour below 2-5°C (check mix design)
- Asphalt: cannot lay below 5°C
- Most coatings/paints: minimum 5-10°C surface temperature
- Roofing membranes: check manufacturer's minimum temperature

### Wind Constraints
- Tower crane: typically cease at 35-45 mph (check crane chart)
- Mobile crane: reduced capacity in wind > 20 mph
- Cladding installation: cease at 25-30 mph
- Working at height: risk assess at > 20 mph

### Seasonal Planning (Northern Hemisphere)
- **Oct-Mar**: Reduced daylight, cold weather, frost risk. Plan indoor/sheltered work.
- **Apr-Sep**: Best period for external works, concrete, steelwork, roofing
- **Nov-Feb**: Highest weather disruption risk. Include weather days in programme.
- **Typical allowance**: 2-5 weather days per month (winter), 0-2 days (summer)
