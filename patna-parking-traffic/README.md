# Parking & Traffic Management: Issues and Challenges (Patna)

A 15-slide presentation prepared for a session with **City Managers in Patna** on urban
parking and traffic management.

## Contents

| File | Description |
|------|-------------|
| `Parking_Traffic_Management_Patna.pptx` | The final 15-slide presentation |
| `generate_parking_ppt.py` | Python script that generates the PPT (uses `python-pptx`) |
| `make_traffic_images.py` | Helper script that renders illustration images |
| `ppt_images/` | Images embedded in the deck |

## Slide Outline

1. Title
2. Session Agenda
3. Why Parking & Traffic Matter
4. The Scale of the Problem in Patna
5. Key Parking Issues in Patna
6. Traffic Management Challenges in Patna
7. Root Causes Behind the Chaos
8. Impact of Poor Management
9. Solutions: Smarter Parking
10. Solutions: Traffic Management
11. The Role of Smart Technology
12. Case Study: Pune Parking Reform (ITDP India)
13. Action Roadmap for Patna
14. Key Takeaways
15. Thank You

## Key Patna Data (verified)

- 12.5 lakh+ vehicles on Patna roads; only ~798 traffic cops
- 62%+ of newly registered vehicles are two-wheelers
- Only ~8% of city area is road (vs 25% ideal)
- Choke points: Boring Road, Gandhi Maidan, Station Road, Chirayatand Pul
- Late-2025: Patna traffic police banned car/auto parking on
  Income Tax Golambar–Dak Bungalow & Budh Marg–GPO Golambar

Sources: Times of India (Patna) reports; NIUA public perception study.

## Regenerate the PPT

```bash
pip install python-pptx pillow matplotlib
python generate_parking_ppt.py
```
