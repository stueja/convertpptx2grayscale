#!/bin/bash

echo .

# mandatory programs
for NEEDED in unzip mogrify zip
do
    command -v $NEEDED >/dev/null 2>&1
    if [[ $? -ne 0 ]]
    then
        echo "$NEEDED not installed"
        exit 4
    fi
done

USAGE(){
    echo "converts png images in pptx files to grayscale"
    echo
    echo "Syntax: $0 [-h] filename"
    echo "Options: "
    echo "d     to be created target directory"
    echo "h     help"
}


# geht mit positional parameters, z. B. script.sh param1 -d abc param2
# https://stackoverflow.com/a/63421397
script_args=()
while [ $OPTIND -le "$#" ]
do
    if getopts d:h options
    then
        case $options
        in
            d) DESTINATION="$OPTARG";;
            h) USAGE;;
        esac
    else
        script_args+=("${!OPTIND}")
        ((OPTIND++))
    fi
done

if [[ "$DESTINATION" == "" ]]
then
    echo "setting DESTINATION to 'grayscale'"
    DESTINATION=grayscale
else
    echo "setting DESTINATION to ${DESTINATION}"
fi


if [[ ${#script_args[@]} -ne 1 ]]
then
    echo "exactly one positional parameter expected (filename), ${#script_args[@]} given."
    echo "${script_args[*]}"
    exit 3
else
    PPTX=$script_args
    echo "Original File: $PPTX"
fi

if [[ ! -e $PPTX ]]
then
    echo "file $PPTX not found"
    exit
fi

PPTXFILE="${PPTX%.*}"
PPTXEXTENSION="${PPTX##*.}"
# modifier for new file name
# MODIFIER="gy"

echo "Making temp directory"
TEMPDIR=$(mktemp -d "/tmp/pptx2gray.XXXXXXXXX")
if [[ $? -ne 0 ]]
then
    echo "error creating temporary directory"
    exit
else
    echo "created temporary directory ${TEMPDIR}"
fi

if [[ ! -d "$TEMPDIR" ]]
then
    echo "error creating $TEMPDIR"
    exit
fi


if [[ ! -w "$TEMPDIR" ]]
then
    echo "$TEMPDIR not writeable"
    exit
fi

echo "unzipping $PPTX to $TEMPDIR"
unzip -q "$PPTX" -d "$TEMPDIR"
if [[ $? -ne 0 ]]
then
    echo "error unzipping $PPTX to $TEMPDIR"
    exit
fi

echo "changing to grayscale"
#mogrify -colorspace gray "$TEMPDIR"/ppt/media/*.png
# convert images except image1.* (title slide background)
find "$TEMPDIR"/ppt/media/ -type f ! -name "image1.*" -exec mogrify -colorspace gray {} \;
if [[ $? -ne 0 ]]
then
    echo "error changing to grayscale"
    exit
fi

echo "creating new pptx"
# need to chdir here, because zip will otherwise
# include all the parent folder names
# -r: recursive
# -o: set pptx timestamp
# -q: quiet
ORIGDIR=$(pwd)
cd "$TEMPDIR"
zip -r -o -q "${PPTXFILE}.${MODIFIER}.${PPTXEXTENSION}" *
if [[ $? -ne 0 ]]
then
    echo "error creating new pptx"
    exit
fi
cd "${ORIGDIR}"

echo "creating destination folder $DESTINATION"
mkdir -p "$DESTINATION"
if [[ $? -ne 0 ]]
then
    echo "error creating destination folder $DESTINATION"
    exit
fi

echo "moving new pptx to destination folder"
# TODO: save output in (configurable) folder instead of
# same folder with different file name
# mv "${PPTXFILE}.${MODIFIER}.${PPTXEXTENSION}" "${ORIGDIR}"
mv "${TEMPDIR}/${PPTXFILE}.${MODIFIER}.${PPTXEXTENSION}" "${DESTINATION}/${PPTXFILE}.${PPTXEXTENSION}"
if [[ $? -ne 0 ]]
then
    echo "error moving new pptx to destination folder"
    exit
fi

echo "removing $TEMPDIR"
rm -rf "$TEMPDIR"
if [[ $? -ne 0 ]]
then
    echo "error removing $TEMPDIR"
    exit
fi

# for some reason, macos (unzip? zip?) additionally creates 
# a folder in $TMPDIR
echo "removing $TMPDIR/pptx2gray*"
rm -rf "$TMPDIR/pptx2gray*"
if [[ $? -ne 0 ]]
then
    echo "error removing $TMPDIR/pptx2gray*"
    exit
fi
