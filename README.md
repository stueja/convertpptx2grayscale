# convertpptx2grayscale
Converts images in a Microsoft Powerpoint pptx to grayscale

Shell script.

Uses `unzip` to extract the contents of the pptx file to a temporary
folder. Uses imagemagick's `mogrify` to convert images to grayscale.
Uses `zip` to compress the files into a new pptx file. 
