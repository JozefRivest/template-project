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
├── indent.log
├── issues
├── makefile
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

## makefile

The template now comes with a `makefile`. Of course, the file is empty. A good 
practice will be to update is soon as possible to keep track of the different 
dependencies in the project, and to ease the process in submitting the replication
package. `makefile` works as follow:

```makefile
target: dependency
    steps to build the target
```

For example

```makefile
.PHONY: paper/0_main.pdf clean

paper/0_main.pdf: paper/0_main.tex \ 
    paper/1_introduction.tex \
    paper/2_theory.tex \
    paper/3_method.tex \
    paper/4_analyses.tex \
    paper/5_conclusion.tex \
    references/references.bib
    cd paper && pdflatex 0_main.tex

clean: 
    cd paper && latexmk -C
```

Once the `makefile` exists, we can run in the terminal different commands to 
either build the entire project easily, or to build specific target such as figures,
tables, or the paper. For example, using the example above, we could run 

```bash
make paper/0_main.pdf
```

to build the paper.
