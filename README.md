# ofc_mediax

Have you ever embedded images into a Powerpoint presentation, and then misplaced or deleted the original images? This is a tool to recover those embedded media files.

ofc_mediax extracts media files (such as images and movies) from pptx, docx, xlsx, and related files. Limited support for OpenDocument formats.

### Usage Examples

To extract all media files from a pptx file named `slides.pptx` into a new directory called `outdir`:
```
ofc_mediax slides.pptx outdir
```

To extract only files with a certain extension (in this example, jpg and png files):
```
ofc_mediax -x .jpg -x .png document.docx outdir
```

To list embedded media files, but not extract them: 
```
ofc_mediax -l slides.pptx
```

### Installation

Either clone this repository with
```
git clone https://github.com/johnhcc/ofc_mediax.git
```
or [download](https://github.com/johnhcc/ofc_mediax/releases) a zip or tarball and extract the contents.

Then simply place the `ofc_mediax` file somewhere in your path.

Requires either Python 2 or 3. Run it by simply typing `ofc_mediax` (or `python ofc_mediax` if you prefer, or `python.exe ofc_mediax` on Windows).

Run without any arguments for usage.

### Contact information:

John M. Haynes<br/>
Website: https://github.com/johnhcc<br/>
Email: johnhcc@vorticity.cc
