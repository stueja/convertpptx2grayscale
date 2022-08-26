#!/bin/bash


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
    echo "h     help"
}


# geht mit positional parameters, z. B. script.sh param1 -d abc param2
# https://stackoverflow.com/a/63421397
script_args=()
while [ $OPTIND -le "$#" ]
do
    if getopts h options
    then
        case $options
        in
            h) USAGE;;
        esac
    else
        script_args+=("${!OPTIND}")
        ((OPTIND++))
    fi
done


if [[ ${#script_args[@]} -ne 1 ]]
then
    echo "exactly one positional parameter expected (filename), ${#script_args[@]} given."
    exit 3
else
    PPTX=$script_args
    echo "Original File: $PPTX"
    echo .
fi

if [[ ! -e $PPTX ]]
then
    echo "file $PPTX not found"
    exit
fi

PPTXFILE="${PPTX%.*}"
PPTXEXTENSION="${PPTX##*.}"

# echo "Making temp directory"
TEMPDIR=$(mktemp -d )
if [[ $? -ne 0 ]]
then
    echo "error creating temporary directory"
    exit
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
mogrify -colorspace gray "$TEMPDIR"/ppt/media/*.png
if [[ $? -ne 0 ]]
then
    echo "error changing to grayscale"
    exit
fi

echo "creating new pptx"
# need a deviation here, because zip will otherwise
# include all the parent folder names
# -r: recursive
# -o: set pptx timestamp
# -q: quiet
NOWDIR=$(pwd)
cd "$TEMPDIR"
zip -r -o -q "${PPTXFILE}.gy.${PPTXEXTENSION}" *
if [[ $? -ne 0 ]]
then
    echo "error creating new pptx"
    exit
fi

echo "moving new pptx to source folder"
mv "${PPTXFILE}.new.${PPTXEXTENSION}" "${NOWDIR}"
if [[ $? -ne 0 ]]
then
    echo "error moving new pptx to source folder"
    exit
fi
cd "${NOWDIR}"

echo "removing $TEMPDIR"
rm -rf "$TEMPDIR"
if [[ $? -ne 0 ]]
then
    echo "error removing $TEMPDIR"
    exit
fi
