"""Deterministic sample dataset for the Milestone Trend Analysis workbook.

Each milestone has a baseline date and a set of forecast changes keyed by
reporting-period index.  The simulation walks through each project's reporting
calendar, applies the changes and records one observation per reporting date.
When a forecast falls on or before a reporting date, the milestone counts as
completed (actual = forecast) and gets no further observations.
"""
from datetime import date, timedelta

D = date


def every(start, days, n):
    return [start + timedelta(days=days * i) for i in range(n)]


PROJECTS = [
    dict(id="PRJ-001", name="ERP Cloud Migration", pm="Sarah Mitchell", sponsor="Chief Financial Officer",
         dept="Information Technology", start=D(2026, 5, 18), finish=D(2027, 2, 26), freq="Bi-weekly",
         desc="Migrate on-premise ERP to a cloud platform, including data migration and cutover.",
         reports=every(D(2026, 6, 1), 14, 13)),
    dict(id="PRJ-002", name="Regional Distribution Centre", pm="David Okafor", sponsor="Chief Operating Officer",
         dept="Operations", start=D(2025, 12, 1), finish=D(2027, 3, 31), freq="Monthly",
         desc="Design and construct a 40,000 m2 automated distribution centre.",
         reports=[D(2026, 1, 12), D(2026, 2, 9), D(2026, 3, 9), D(2026, 4, 13), D(2026, 5, 11), D(2026, 6, 8),
                  D(2026, 7, 13), D(2026, 8, 10), D(2026, 9, 14), D(2026, 10, 12), D(2026, 11, 9)]),
    dict(id="PRJ-003", name="Customer Portal Redesign", pm="Priya Nair", sponsor="Chief Digital Officer",
         dept="Digital", start=D(2026, 4, 20), finish=D(2027, 1, 29), freq="Every 3 weeks",
         desc="Replace the customer self-service portal with a new design system and payments module.",
         reports=every(D(2026, 5, 4), 21, 10)),
]

CATEGORIES = ["Governance", "Design", "Procurement", "Build", "Construction", "Testing", "Deployment", "Handover"]
PRIORITIES = ["Critical", "High", "Medium", "Low"]
DEPARTMENTS = ["Information Technology", "Operations", "Digital", "Finance", "Engineering", "Procurement", "PMO"]

# (id, name, description, category, baseline, owner, department, priority, {period: delta}, init_offset, skip_periods)
MILESTONES = {
    "PRJ-001": [
        ("M-101", "Project Charter Approved", "Charter and business case signed off by steering committee", "Governance", D(2026, 6, 12), "Sarah Mitchell", "PMO", "High", {}, 0, []),
        ("M-102", "Requirements Sign-off", "Functional and non-functional requirements baselined", "Design", D(2026, 7, 3), "Tom Becker", "Information Technology", "High", {1: 5, 2: 4}, 0, []),
        ("M-103", "Solution Architecture Approved", "Target architecture approved by design authority", "Design", D(2026, 7, 24), "Lena Fischer", "Information Technology", "High", {2: -4}, 0, []),
        ("M-104", "Data Migration Strategy", "Migration approach, tooling and reconciliation rules agreed", "Design", D(2026, 8, 14), "Raj Malhotra", "Information Technology", "Medium", {}, 0, []),
        ("M-105", "Cloud Environment Provisioned", "Production and non-production tenants provisioned", "Build", D(2026, 8, 28), "Lena Fischer", "Information Technology", "Critical", {4: 21}, 0, []),
        ("M-106", "Integration Build Complete", "All interfaces built and unit tested", "Build", D(2026, 10, 2), "Tom Becker", "Information Technology", "Critical", {4: 7, 6: 7, 8: 7, 10: 5}, 0, []),
        ("M-107", "Data Migration Trial 1", "First full trial load with reconciliation", "Testing", D(2026, 10, 23), "Raj Malhotra", "Information Technology", "High", {5: 7, 8: 10}, 0, []),
        ("M-108", "System Integration Testing Complete", "SIT exit criteria met", "Testing", D(2026, 11, 27), "Anna Kowalski", "Information Technology", "Critical", {4: 7, 6: 7, 8: 7, 10: 7, 12: 7}, 0, []),
        ("M-109", "User Acceptance Testing Complete", "Business sign-off of UAT", "Testing", D(2026, 12, 18), "Anna Kowalski", "Finance", "Critical", {6: 7, 10: 14}, 0, []),
        ("M-110", "Training Delivered", "End-user training delivered to 450 users", "Deployment", D(2027, 1, 8), "Maria Lopez", "Finance", "Medium", {}, 0, []),
        ("M-111", "Cutover Rehearsal", "Dress rehearsal of cutover runbook", "Deployment", D(2027, 1, 22), "Tom Becker", "Information Technology", "High", {10: 7}, 0, []),
        ("M-112", "Go-Live", "Production go-live", "Deployment", D(2027, 2, 5), "Sarah Mitchell", "PMO", "Critical", {10: 7, 12: 7}, 0, []),
        ("M-113", "Hypercare Exit", "Hypercare exit criteria met", "Deployment", D(2027, 2, 26), "Maria Lopez", "Information Technology", "Medium", {10: 7, 12: 7}, 0, []),
        ("M-114", "Legacy Decommissioning", "Legacy ERP archived and switched off", "Governance", D(2027, 3, 19), "Raj Malhotra", "Information Technology", "Low", {7: -14}, 0, [12]),
    ],
    "PRJ-002": [
        ("M-201", "Planning Permission Granted", "Full planning consent from local authority", "Governance", D(2026, 2, 20), "David Okafor", "Operations", "Critical", {1: 14}, 0, []),
        ("M-202", "Site Acquisition Complete", "Land purchase completed", "Procurement", D(2026, 3, 13), "Helen Price", "Finance", "Critical", {}, 0, []),
        ("M-203", "Design Freeze", "RIBA stage 4 design frozen", "Design", D(2026, 4, 17), "Marco Rossi", "Engineering", "High", {2: -7}, 0, []),
        ("M-204", "Main Contractor Awarded", "Design-and-build contract signed", "Procurement", D(2026, 5, 29), "Helen Price", "Procurement", "High", {2: 10, 4: 5}, 0, []),
        ("M-205", "Groundworks Complete", "Foundations, drainage and slab complete", "Construction", D(2026, 7, 31), "Marco Rossi", "Engineering", "High", {4: 14, 6: 7}, 0, []),
        ("M-206", "Steel Frame Erected", "Structural steel frame erected and signed off", "Construction", D(2026, 9, 25), "Marco Rossi", "Engineering", "High", {4: 21, 6: 10, 8: -5}, 0, []),
        ("M-207", "Building Watertight", "Roof and cladding complete", "Construction", D(2026, 11, 20), "Marco Rossi", "Engineering", "Critical", {5: 14, 7: 10, 9: 7}, 0, []),
        ("M-208", "MEP Installation Complete", "Mechanical, electrical and plumbing installed", "Construction", D(2027, 1, 15), "James Carter", "Engineering", "High", {5: 10, 8: 14, 10: 7}, 0, [6]),
        ("M-209", "Racking & Automation Installed", "Racking, conveyors and sortation commissioned", "Build", D(2027, 2, 5), "James Carter", "Operations", "High", {7: -3, 10: -5}, 0, []),
        ("M-210", "WMS Integration Tested", "Warehouse management system integration tested", "Testing", D(2027, 2, 19), "Nina Patel", "Information Technology", "Medium", {}, 0, []),
        ("M-211", "Practical Completion", "Building handed over by main contractor", "Handover", D(2027, 3, 5), "David Okafor", "Engineering", "Critical", {6: 14, 9: 14}, 0, []),
        ("M-212", "Operational Readiness Review", "Readiness gate passed", "Governance", D(2027, 3, 19), "Nina Patel", "Operations", "Medium", {9: 7, 10: -4}, 0, []),
        ("M-213", "Site Go-Live", "First live outbound shipment", "Handover", D(2027, 3, 31), "David Okafor", "Operations", "Critical", {6: 7, 9: 14}, 0, []),
    ],
    "PRJ-003": [
        ("M-301", "Discovery Complete", "Discovery findings and roadmap agreed", "Design", D(2026, 5, 22), "Priya Nair", "Digital", "Medium", {}, 0, []),
        ("M-302", "UX Research Findings", "User research synthesis published", "Design", D(2026, 6, 12), "Owen Hughes", "Digital", "Medium", {1: 3}, 0, []),
        ("M-303", "Design System v1", "Component library v1 released", "Design", D(2026, 7, 10), "Owen Hughes", "Digital", "High", {2: -5}, 0, []),
        ("M-304", "Prototype Usability Tested", "Clickable prototype tested with 30 customers", "Testing", D(2026, 7, 31), "Owen Hughes", "Digital", "Medium", {}, 0, []),
        ("M-305", "API Gateway Ready", "API gateway and auth service in production", "Build", D(2026, 8, 28), "Kenji Watanabe", "Information Technology", "High", {3: -7}, 0, []),
        ("M-306", "Account Module Build", "Account management features complete", "Build", D(2026, 9, 25), "Kenji Watanabe", "Digital", "High", {4: 7}, 0, []),
        ("M-307", "Payments Module Build", "Payments and billing features complete", "Build", D(2026, 10, 23), "Kenji Watanabe", "Digital", "Critical", {3: 14, 5: 7, 7: 7}, 0, []),
        ("M-308", "Accessibility Audit", "WCAG 2.2 AA audit passed", "Testing", D(2026, 11, 20), "Grace Liu", "Digital", "High", {}, 0, []),
        ("M-309", "Performance Testing", "Load test at 3x peak passed", "Testing", D(2026, 12, 11), "Grace Liu", "Information Technology", "High", {6: 10}, 0, []),
        ("M-310", "Beta Launch", "Beta opened to 5% of customers", "Deployment", D(2026, 12, 18), "Priya Nair", "Digital", "High", {5: -4}, 0, []),
        ("M-311", "Content Migration Complete", "All help content migrated", "Build", D(2027, 1, 8), "Grace Liu", "Digital", "Low", {}, 0, []),
        ("M-312", "Public Launch", "Portal launched to all customers", "Deployment", D(2027, 1, 29), "Priya Nair", "Digital", "Critical", {6: 7, 9: 7}, 0, []),
        ("M-313", "Post-Launch Benefits Review", "Benefits realisation review", "Governance", D(2027, 4, 30), "Priya Nair", "PMO", "Low", {}, 0, list(range(10))),
    ],
}

REASONS_LATE = ["Vendor delivery delay", "Dependency slipped", "Resource constraint", "Scope change approved",
                "Defects found in testing", "Weather impact on site", "Approval cycle extended"]
REASONS_EARLY = ["Dependency delivered early", "Additional resources assigned", "Scope simplified",
                 "Parallel workstream accelerated"]


def build():
    """Return (projects, milestones, observations)."""
    milestones, obs = [], []
    for p in PROJECTS:
        for (mid, name, desc, cat, base, owner, dept, prio, deltas, init, skip) in MILESTONES[p["id"]]:
            milestones.append(dict(project=p["id"], id=mid, name=name, desc=desc, cat=cat, base=base,
                                   owner=owner, dept=dept, prio=prio))
            fc = base + timedelta(days=init)
            done = False
            for i, r in enumerate(p["reports"]):
                if done:
                    break
                d = deltas.get(i, 0)
                fc = fc + timedelta(days=d)
                if i in skip:
                    continue
                actual = None
                comment = ""
                if d > 0:
                    comment = REASONS_LATE[(i + int(mid[-2:])) % len(REASONS_LATE)] + f" (+{d} d)"
                elif d < 0:
                    comment = REASONS_EARLY[(i + int(mid[-2:])) % len(REASONS_EARLY)] + f" ({d} d)"
                if fc <= r:
                    actual = fc
                    done = True
                    comment = "Milestone achieved" + (f"; {comment}" if comment else "")
                obs.append(dict(mid=mid, rep=r, fc=fc, actual=actual, comment=comment))
    return PROJECTS, milestones, obs


if __name__ == "__main__":
    _, ms, ob = build()
    print(len(ms), "milestones", len(ob), "observations")
