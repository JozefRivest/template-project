$aux_dir="build";
$pdflatex = 'texcount -inc -sum -1 %S > wordcount.tex; pdflatex -interaction=nonstopmode -halt-on-error %O %S';
