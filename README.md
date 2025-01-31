# PPTX Editor
A dumb script that uses PowerPoint to create projects in C.

## Dependencies
- Python 3
- python-pptx
- clang (you can edit the script to use another compiler)

## How to use
1. Put a pptx file into this directory with the default template.
Make sure to only use the basic "title" and "title and content" slides.

2. Put the name of your project in the title box of a "title" slide.
This should be the first slide.

3. The rest of the slides should be "title and content" slides.
The titles are the names of the files and the contents are the code.

4. Run the `compile.py` script, and reload the pptx file.
The output will be in the subtitle box of the title slide.
