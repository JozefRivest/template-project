# Project

Brief description of the project and its structure.

```
.
├── code
│   ├── 01_cleaning.R
│   ├── 02_analysis.R
│   ├── 03_figures.R
│   └── 04_tables.R
├── data
│   ├── processed
│   └── raw
├── issues
├── notes
│   ├── org_table.py
│   └── readings.yaml
├── paper
│   ├── 0_main.tex
│   ├── 1_introduction.tex
│   ├── 2_literature.tex
│   ├── 3_method.tex
│   ├── 4_analyses.tex
│   ├── 5_conclusion.tex
│   ├── figures
│   └── tables
├── presentation
│   └── presentation.tex
├── README.md
└── references
    └── references.bib
```

## GitHub issues

If you wish to edit your issues within your terminal, without going to
GitHub website everytime, you can add the following function to your 
.bashrc or .zshrc

```bash
get_issue() {
  local issue="$1"
  gh issue view "${issue}" --json body --jq '.body' > issues/"issue_${issue}.md"
}
```

With this function, you can fetch issues form GitHub directly into a markdown
file that will be created in the `issues` directory. Once you make the changes
to the issue's text, you then run `gh issue edit $$ body-file issues/issue_$$.md`
where the dollar signs stand for the number of the issue. This will update the 
issue on GitHub.

Everybody is welcome to use this template to organize their working 
environment. You can fork or clone the repo as you wish.
