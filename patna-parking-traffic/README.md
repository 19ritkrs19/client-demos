# Parking & Traffic Management: Issues and Challenges (Patna) — v2

A 15-slide presentation for a session with **City Managers in Patna**.

**Session by Pourush Agarwal, DySP Traffic Patna.**

## Contents

| File | Description |
|------|-------------|
| `Parking_Traffic_Management_Patna_v2.pptx` | Final 15-slide presentation |
| `generate_parking_ppt.py` | Script that generates the base deck (`python-pptx`) |
| `patch_slides.py` | Script that swaps in specific photos on individual slides |
| `fix_alignment.py` | Script that fixes hanging-indent alignment on wrapped bullets |
| `ppt_images/` | Real Patna photos + Pune case-study photo embedded in the deck |

## Slide Outline

1. Title (DySP credit + Patna Traffic Police cover photo)
2. Session Agenda
3. Why Parking & Traffic Matter
4. Key Parking Issues in Patna (No Parking board + Wakarganj photos)
5. Traffic Management Challenges (junction regulation photo)
6. Special Traffic Situations — VIP Movement & Metro Diversion (Boring Road photo)
7. Root Causes Behind the Chaos
8. Impact of Poor Management (congestion photo)
9. Solutions: Smarter Parking (organised auto-stand photo)
10. Solutions: Traffic Management (sign boards, real-time updates, NCC, flyovers, ring road, ICCC)
11. The Role of Smart Technology
12. Case Study: Pune Parking Reform (ITDP India, with photo)
13. Action Roadmap for Patna
14. Key Takeaways
15. Thank You (enhanced, sharpened photo)

## Topics Covered

Core parking & traffic issues, plus:
- **VIP movement** — route clearance & advance planning
- **Metro diversion** — lane closures & planned diversions
- **Sign boards** — signage placement & enforcement
- **Real-time updates** — live traffic info to citizens
- **Elevated corridors** — Karbigahiya–Mithapur Flyover, Mithapur–Sipara Elevated Road
- **Ring road development**, **Smart management (ICCC)**, **community awareness by NCC**

## Key Patna Data (verified)

- 12.5 lakh+ vehicles; only ~798 traffic cops
- 62%+ of new vehicles are two-wheelers
- Only ~8% of city area is road (vs 25% ideal)
- Choke points: Boring Road, Gandhi Maidan, Station Road
- Late-2025: Patna traffic police banned car/auto parking on key stretches

Sources: Times of India (Patna) reports; NIUA public perception study.

## Regenerate / re-patch the PPT

```bash
pip install python-pptx pillow
python generate_parking_ppt.py   # rebuild base deck
python patch_slides.py           # apply specific photo swaps
python fix_alignment.py          # fix bullet hanging-indent alignment
```
