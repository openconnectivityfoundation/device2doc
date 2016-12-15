#!/bin/bash

PYTHON_EXE=C:\\python35-32\\python3.exe
RAML2DOC=../src/device2doc.py

OUTPUT_DIR=./out
OUTPUT_DIR_DOCS=../test/$OUTPUT_DIR
REF_DIR=./ref
EXT=.txt

function compare_output {
    diff -w $OUTPUT_DIR/$TEST_CASE$EXT $REF_DIR/$TEST_CASE$EXT
    echo "testcase difference: $TEST_CASE $?"
    #echo "blah"
}

function compare_to_reference_file {
    diff -w $OUTPUT_DIR/$1 $REF_DIR/$1
    echo "output $1 difference: $TEST_CASE $?"
    #echo "blah"
}


function compare_to_reference_file_in_dir {
    diff -w $OUTPUT_DIR/$1 $REF_DIR/$2/$1
    echo "output $1 difference: $TEST_CASE $?"
    #echo "blah"
}

function compare_file {
    echo "comparing ($TEST_CASE): " $1 $2
    diff -wb $1 $2
    #echo "blah"
}


function my_test {
    $PYTHON_EXE $RAML2DOC $* > $OUTPUT_DIR/$TEST_CASE$EXT 2>&1
    compare_output
} 

function my_test_in_dir {
    mkdir -p $OUTPUT_DIR/$TEST_CASE
    $PYTHON_EXE $RAML2DOC $* > $OUTPUT_DIR/$TEST_CASE/$TEST_CASE$EXT 2>&1
    compare_file $OUTPUT_DIR/$TEST_CASE/$TEST_CASE$EXT $REF_DIR/$TEST_CASE/$TEST_CASE$EXT
} 


TEST_CASE="testcase_1"

function tests {

# option -h
TEST_CASE="testcase_1"
my_test -h

# default docx 
TEST_CASE="testcase_2"
my_test_in_dir -docx ../input/ResourceTemplate.docx -device ../test/in/test_1/device.json -word_out $OUTPUT_DIR_DOCS/$TEST_CASE/$TEST_CASE.docx

# default docx 
TEST_CASE="testcase_3"
my_test_in_dir -docx ../input/ResourceTemplate.docx -lbnldevice ../test/in/test_2/device.json -word_out $OUTPUT_DIR_DOCS/$TEST_CASE/$TEST_CASE.docx

}

tests  
