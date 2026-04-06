# GAT Student Reports

Generates PIN-locked HTML diagnostic reports for GAT prep students.

## Quick start

```bash
pip install openpyxl
python generate_reports.py \
    --question-bank ../GAT_Question_Bank.xlsx \
    --responses ../response.csv \
    --output-dir docs/ \
    --logo-path ../noon\ logo\ v2.png
```

Reports land in `docs/`. Push the repo and enable GitHub Pages from `docs/` on the `main` branch.

Each student gets `docs/<student-slug>.html`, PIN-locked with the last 4 digits of their phone number.
